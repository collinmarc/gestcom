Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Authentication
Imports System.IO
Imports System.Net.Sockets
Imports System.Net.Security
Imports System.Net
Imports System.Collections
Imports System

Namespace ImapVB
    Public Class ImapBase
        ' IMAP Message Flags
        ''' <summary>
        ''' Message flag:seen
        ''' </summary>
        Protected Const IMAP_MSG_FLAG_SEEN As String = "seen"
        ''' <summary>
        ''' Message flag: answered
        ''' </summary>
        Protected Const IMAP_MSG_FLAG_ANSWERED As String = "Answered"
        ''' <summary>
        ''' Message flag: flagged
        ''' </summary>
        Protected Const IMAP_MSG_FLAG_FLAGGED As String = "flagged"
        ''' <summary>
        ''' Message flag: draft
        ''' </summary>
        Protected Const IMAP_MSG_FLAG_DRAFT As String = "draft"
        ''' <summary>
        ''' Message flag: deleted
        ''' </summary>
        Protected Const IMAP_MSG_FLAG_DELETED As String = "deleted"
        ''' <summary>
        ''' Max Message flags: 10
        ''' </summary>
        Protected Const IMAP_MAX_MSG_FLAGS As UShort = 10
        ''' <summary>
        ''' Imap default port: 143
        ''' </summary>
        Protected Const IMAP_DEFAULT_PORT As UShort = 143
        ''' <summary>
        ''' Imap default timeout:30 sec
        ''' </summary>
        Protected Const IMAP_DEFAULT_TIMEOUT As UShort = 30
        ''' <summary>
        ''' Imap Command Identifier value:Initial 0
        ''' </summary>
        Protected Shared IMAP_COMMAND_VAL As UShort = 0
        ''' <summary>
        ''' Imap command Identified prefix: IMAP00
        ''' </summary>
        Protected Const IMAP_COMMAND_PREFIX As String = "IMAP00"
        ''' <summary>
        ''' Imap command identified which is combination of
        ''' Imap identifier prefix and val
        ''' eg. Prefix:IMAP00, Val: 1
        ''' Imap command Identified= IMAP001
        ''' </summary>
        Protected ReadOnly Property IMAP_COMMAND_IDENTIFIER() As String
            Get
                Return IMAP_COMMAND_PREFIX + IMAP_COMMAND_VAL.ToString() + " "
            End Get
        End Property

        ''' <summary>
        ''' Imap Server OK response which is combination of 
        ''' Imap Identifier and Imap OK response.
        ''' eg. IMAP001 OK
        ''' </summary>
        Protected ReadOnly Property IMAP_SERVER_RESPONSE_OK() As String
            Get
                Return IMAP_COMMAND_IDENTIFIER + IMAP_OK_RESPONSE
            End Get
        End Property
        ''' <summary>
        ''' Imap Server NO response which is combination of 
        ''' Imap Identifier and Imap NO response.
        ''' eg. IMAP001 NO
        ''' </summary>
        Protected ReadOnly Property IMAP_SERVER_RESPONSE_NO() As String
            Get
                Return IMAP_COMMAND_IDENTIFIER + IMAP_NO_RESPONSE
            End Get
        End Property
        ''' <summary>
        ''' Imap Server BAD response which is combination of
        ''' Imap Identifier and Imap BAD response.
        ''' eg. IMAP001 BAD
        ''' </summary>
        Protected ReadOnly Property IMAP_SERVER_RESPONSE_BAD() As String
            Get
                Return IMAP_COMMAND_IDENTIFIER + IMAP_BAD_RESPONSE
            End Get
        End Property

        ''' <summary>
        ''' Imap Untagged response prefix: *
        ''' </summary>
        Protected Const IMAP_UNTAGGED_RESPONSE_PREFIX As String = "*"
        ''' <summary>
        ''' Impa ok response : OK
        ''' </summary>
        Protected Const IMAP_OK_RESPONSE As String = "OK"
        ''' <summary>
        ''' Imap no response: NO
        ''' </summary>
        Protected Const IMAP_NO_RESPONSE As String = "NO"
        ''' <summary>
        ''' Imap bad response :BAD
        ''' </summary>
        Protected Const IMAP_BAD_RESPONSE As String = "BAD"
        ''' <summary>
        ''' Imap bad server response : "* BAD"
        ''' </summary>
        Protected Const IMAP_BAD_SERVER_RESPONSE As String = "* BAD"
        ''' <summary>
        ''' Imap ok server response: "* OK"
        ''' </summary>
        Protected Const IMAP_OK_SERVER_RESPONSE As String = "* OK"

        ''' <summary>
        ''' Imap Server response "* CAPABILITY"
        ''' </summary>
        Protected Const IMAP_CAPABILITY_SERVER_RESPONSE As String = "* CAPABILITY"
        ''' <summary>
        ''' Imap capability command : CAPABILITY
        ''' </summary>
        Protected Const IMAP_CAPABILITY_COMMAND As String = "CAPABILITY"
        ''' <summary>
        ''' Imap connect command :CONNECT
        ''' </summary>
        Protected Const IMAP_CONNECT_COMMAND As String = "CONNECT"
        ''' <summary>
        ''' Imap login command : LOGIN userid  password
        ''' </summary>
        Protected Const IMAP_LOGIN_COMMAND As String = "LOGIN"
        ''' <summary>
        ''' Imap logout command : LOGOUT
        ''' </summary>
        Protected Const IMAP_LOGOUT_COMMAND As String = "LOGOUT"
        ''' <summary>
        ''' Imap select command : SELECT INBOX
        ''' </summary>
        Protected Const IMAP_SELECT_COMMAND As String = "SELECT"
        ''' <summary>
        ''' Imap examine command : EXAMINE INBOX
        ''' </summary>
        Protected Const IMAP_EXAMINE_COMMAND As String = "EXAMINE"
        ''' <summary>
        ''' Imap append command : APPEND
        ''' </summary>
        Protected Const IMAP_APPEND_COMMAND As String = "APPEND"
        ''' <summary>
        ''' Imap quota response : QUOTA
        ''' </summary>
        Protected Const IMAP_QUOTA_RESPONSE As String = "QUOTA"
        ''' <summary>
        ''' Imap get quota command : GETQUOTAROOT 
        ''' </summary>
        Protected Const IMAP_GETQUOTA_COMMAND As String = "GETQUOTAROOT"
        ''' <summary>
        ''' Imap append response start : [
        ''' </summary>
        Protected Const IMAP_APPEND_RESPONSE_START As Char = "["c
        ''' <summary>
        ''' Imap append response end : ]
        ''' </summary>
        Protected Const IMAP_APPEND_RESPONSE_END As Char = "]"c
        ''' <summary>
        ''' Imap go ahead response: +
        ''' </summary>
        Protected Const IMAP_GO_AHEAD_RESPONSE As Char = "+"c
        ''' <summary>
        ''' Imap uid search command : UID SEARCH
        ''' </summary>
        Protected Const IMAP_SEARCH_COMMAND As String = "UID SEARCH"
        ''' <summary>
        ''' Imap search command : SEARCH
        ''' </summary>
        Protected Const IMAP_SEARCH_RESPONSE As String = "SEARCH"
        ''' <summary>
        ''' Imap uid fetch command : UID FETCH
        ''' </summary>
        Protected Const IMAP_UIDFETCH_COMMAND As String = "UID FETCH"
        ''' <summary>
        ''' Imap fetch command : FETCH
        ''' </summary>
        Protected Const IMAP_FETCH_COMMAND As String = "FETCH"
        ''' <summary>
        ''' Imap BodyStructure command
        ''' </summary>
        Protected Const IMAP_BODYSTRUCTURE_COMMAND As String = "BODYSTRUCTURE"
        ''' <summary>
        ''' Imap uid store command
        ''' </summary>
        Protected Const IMAP_UIDSTORE_COMMAND As String = "UID STORE"
        Protected Const IMAP_STORE_COMMAND As String = "STORE"

        ''' <summary>
        ''' Imap uid copy command
        ''' </summary>
        Protected Const IMAP_UIDCOPY_COMMAND As String = "UID COPY"
        Protected Const IMAP_COPY_COMMAND As String = "COPY"
        ''' <summary>
        ''' Imap expunge command
        ''' </summary>
        Protected Const IMAP_EXPUNGE_COMMAND As String = "EXPUNGE"
        ''' <summary>
        ''' Imap noop command : NOOP
        ''' </summary>
        Protected Const IMAP_NOOP_COMMAND As String = "NOOP"

        ''' <summary>
        ''' Imap add flags +flags
        ''' </summary>
        Protected Const IMAP_SETFLAGS_COMMAND As String = "+FLAGS"

        ''' <summary>
        ''' Imap remove flags -flags
        ''' </summary>
        Protected Const IMAP_REMOVEFLAGS_COMMAND As String = "-FLAGS"

        ''' <summary>
        ''' Imap RFC822.SIZE
        ''' </summary>
        Protected Const IMAP_RFC822_SIZE_COMMAND As String = "RFC822.SIZE"
        ''' <summary>
        ''' Imap command terminator: \r\n
        ''' </summary>
        Protected Const IMAP_COMMAND_EOL As String = vbCr & vbLf
        ''' <summary>
        ''' Imap message nil size : NIL
        ''' </summary>
        Protected Const IMAP_MESSAGE_NIL As String = "NIL"
        ''' <summary>
        ''' Imap message header terminator : \r\n
        ''' </summary>
        Protected Const IMAP_MESSAGE_HEADER_EOL As String = vbCr & vbLf
        ''' <summary>
        ''' Imap message size start : '{'
        ''' </summary>
        Protected Const IMAP_MESSAGE_SIZE_START As Char = "{"c
        ''' <summary>
        ''' Imap message size end : '}'
        ''' </summary>
        Protected Const IMAP_MESSAGE_SIZE_END As Char = "}"c
        ''' <summary>
        ''' Imap message content type : "Content-Type: "
        ''' </summary>
        Protected Const IMAP_MESSAGE_CONTENT_TYPE As String = "content-type"
        ''' <summary>
        ''' Imap mesage content type: rfc822
        ''' </summary>
        Protected Const IMAP_MESSAGE_RFC822 As String = "message/rfc822"
        ''' <summary>
        ''' Imap message id
        ''' </summary>
        Protected Const IMAP_MESSAGE_ID As String = "message-id"
        ''' <summary>
        ''' Imap mesage content type: multipart
        ''' </summary>
        Protected Const IMAP_MESSAGE_MULTIPART As String = "multipart"
        ''' <summary>
        ''' Imap content encoding : "Content-Transfer-Encoding: "
        ''' </summary>
        Protected Const IMAP_MESSAGE_CONTENT_ENCODING As String = "content-transfer-encoding"
        ''' <summary>
        ''' Imap content description : "Content-Description: "
        ''' </summary>
        Protected Const IMAP_MESSAGE_CONTENT_DESC As String = "content-description"
        ''' <summary>
        ''' Imap content disposition : "Content-Disposition: "
        ''' </summary>
        Protected Const IMAP_MESSAGE_CONTENT_DISP As String = "content-disposition"
        ''' <summary>
        ''' Imap content size
        ''' </summary>
        Protected Const IMAP_MESSAGE_CONTENT_SIZE As String = "content-size"
        ''' <summary>
        ''' Imap content lines
        ''' </summary>
        Protected Const IMAP_MESSAGE_CONTENT_LINES As String = "content-lines"

        ''' <summary>
        ''' Imap message base64 encoding : BASE64
        ''' </summary>
        Protected Const IMAP_MESSAGE_BASE64_ENCODING As String = "base64"
        ''' <summary>
        ''' Imap message default part : 1
        ''' </summary>
        Protected Const IMAP_MSG_DEFAULT_PART As String = "1"
        ''' <summary>
        ''' Imap header Sender tag
        ''' </summary>
        Protected Const IMAP_HEADER_SENDER_TAG As String = "sender"
        ''' <summary>
        ''' Imap header from tag
        ''' </summary>
        Protected Const IMAP_HEADER_FROM_TAG As String = "from"
        ''' <summary>
        ''' Imap header in-reply-to tag
        ''' </summary>
        Protected Const IMAP_HEADER_IN_REPLY_TO_TAG As String = "in-reply-to"
        ''' <summary>
        ''' IKmap header reply-to tag
        ''' </summary>
        Protected Const IMAP_HEADER_REPLY_TO_TAG As String = "reply-to"
        ''' <summary>
        ''' Imap header to tag
        ''' </summary>
        Protected Const IMAP_HEADER_TO_TAG As String = "to"
        ''' <summary>
        ''' Imap header cc tag
        ''' </summary>
        Protected Const IMAP_HEADER_CC_TAG As String = "cc"
        ''' <summary>
        ''' Imap header bcc tag
        ''' </summary>
        Protected Const IMAP_HEADER_BCC_TAG As String = "bcc"
        ''' <summary>
        ''' Imap header subject tag
        ''' </summary>
        Protected Const IMAP_HEADER_SUBJECT_TAG As String = "subject"
        ''' <summary>
        ''' Imap header date tag
        ''' </summary>
        Protected Const IMAP_HEADER_DATE_TAG As String = "date"
        ''' <summary>
        ''' Imap body type
        ''' </summary>
        Protected Const IMAP_PLAIN_TEXT As String = "text/plain"
        ''' <summary>
        ''' Imap audio wave:  audio/wav
        ''' </summary>
        Protected Const IMAP_AUDIO_WAV As String = "audio/wav"
        ''' <summary>
        ''' Imap video mpeg4  : video/mpeg4
        ''' </summary>
        Protected Const IMAP_VIDEO_MPEG4 As String = "video/mpeg4"

        ''' <summary>
        ''' Imap host
        ''' </summary>
        Protected m_sHost As String = ""
        ''' <summary>
        ''' Imap port : default IMAP_DEFAULT_PORT : 143
        ''' </summary>
        Protected m_nPort As UShort = IMAP_DEFAULT_PORT
        ''' <summary>
        ''' User id
        ''' </summary>
        Protected m_sUserId As String = ""
        ''' <summary>
        ''' User Password
        ''' </summary>
        Protected m_sPassword As String = ""
        ''' <summary>
        ''' SSL Enabled
        ''' </summary>
        Protected m_bSSLEnabled As Boolean = False
        ''' <summary>
        ''' Is Imap server connected
        ''' </summary>
        Protected m_bIsConnected As Boolean = False
        ''' <summary>
        ''' Tcpclient object
        ''' </summary>
        Private m_oImapServ As TcpClient
        ''' <summary>
        ''' Network stream object
        ''' </summary>
        Private m_oNetStrm As Stream
        ''' <summary>
        ''' StreamReader object
        ''' </summary>
        Private m_oRdStrm As StreamReader

        ''' <summary>
        ''' Imap server response result
        ''' </summary>
        Protected Enum ImapResponseEnum
            ''' <summary>
            ''' Imap Server responded "OK"
            ''' </summary>
            IMAP_SUCCESS_RESPONSE
            ''' <summary>
            ''' Imap Server responded "NO" or "BAD"
            ''' </summary>
            IMAP_FAILURE_RESPONSE
            ''' <summary>
            ''' Imap Server responded "*"
            ''' </summary>
            IMAP_IGNORE_RESPONSE
        End Enum
        ''' <summary>
        ''' Log type enum
        ''' </summary>
        Protected Enum LogTypeEnum
            ''' <summary>
            ''' Information
            ''' </summary>
            INFO
            ''' <summary>
            ''' Warning
            ''' </summary>
            WARN
            ''' <summary>
            ''' Error
            ''' </summary>
            [ERROR]
            ''' <summary>
            ''' Imap Log information
            ''' </summary>
            IMAP
        End Enum
        ''' <summary>
        ''' Logging function
        ''' </summary>
        ''' <param name="type">Log type;LogTypeEnum</param>
        ''' <param name="log">Log data</param>
        Protected Sub Log(type As LogTypeEnum, log As String)
            Select Case type
                Case LogTypeEnum.INFO
                    Console.WriteLine("Info:" + log)
                    Exit Select
                Case LogTypeEnum.WARN
                    Console.WriteLine("Warning:" + log)
                    Exit Select
                Case LogTypeEnum.[ERROR]
                    Console.WriteLine("Error:" + log)
                    Exit Select
                Case LogTypeEnum.IMAP
                    Console.WriteLine(log)
                    Exit Select

            End Select
        End Sub

        ''' <summary>
        ''' Connect to specified host and port
        ''' </summary>
        ''' <param name="sHost">Imap host</param>
        ''' <param name="nPort">Imap port</param>
        ''' <param name="sslEnabled"> </param>
        ''' <returns>ImapResponseEnum type</returns>
        Protected Function Connect(sHost As String, nPort As UShort, Optional sslEnabled As Boolean = False) As ImapResponseEnum
            IMAP_COMMAND_VAL = 0
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim e_connect As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_CONNECT, sHost)
            Try
                m_oImapServ = New TcpClient(sHost, nPort)
                If sslEnabled Then

                    '                    Dim sslStrm As New SslStream(m_oImapServ.GetStream(), False, AdrressOf ValidateServerCertificate() , Nothing)
                    Dim sslStrm As New SslStream(m_oImapServ.GetStream(), False)
                    'm_SSLStream.ReadTimeout = 30000;

                    Try
                        sslStrm.AuthenticateAsClient(sHost)
                    Catch e As AuthenticationException
                        Console.WriteLine("Exception: {0}", e.Message)
                        If e.InnerException IsNot Nothing Then
                            Console.WriteLine("Inner exception: {0}", e.InnerException.Message)
                        End If
                        Console.WriteLine("Authentication failed - closing the connection.")
                        m_oImapServ.Close()
                        Return ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    End Try

                    m_oNetStrm = CType(sslStrm, Stream)
                    m_oRdStrm = New StreamReader(sslStrm)
                Else
                    m_oNetStrm = m_oImapServ.GetStream()
                    m_oRdStrm = New StreamReader(m_oImapServ.GetStream())
                End If


                Dim sResult As String = m_oRdStrm.ReadLine()
                If (sResult.StartsWith(IMAP_OK_SERVER_RESPONSE) = True) OrElse (sResult.StartsWith(IMAP_CAPABILITY_SERVER_RESPONSE) = True) Then
                    Log(LogTypeEnum.IMAP, sResult)
                    eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
                    Capability()
                Else
                    Log(LogTypeEnum.IMAP, sResult)
                    eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                End If
            Catch
                Throw e_connect
            End Try
            m_sHost = sHost
            m_nPort = nPort
            m_bSSLEnabled = sslEnabled
            Return eImapResponse
        End Function


        ''' <summary>
        ''' Disconnect connection with Imap server
        ''' </summary>
        Protected Sub Disconnect()
            IMAP_COMMAND_VAL = 0
            If m_bIsConnected Then
                If m_oNetStrm IsNot Nothing Then
                    m_oNetStrm.Close()
                End If
                If m_oRdStrm IsNot Nothing Then
                    m_oRdStrm.Close()
                End If
            End If

        End Sub

        ''' <summary>
        ''' Send command to server and retrieve response
        ''' </summary>
        ''' <param name="command">Command to send Imap Server</param>
        ''' <param name="sResultArray">Imap Server response</param>
        ''' <returns>ImapResponseEnum type</returns>
        Protected Function SendAndReceive(command As String, ByRef sResultArray As ArrayList) As ImapResponseEnum
            IMAP_COMMAND_VAL += 1
            Dim sCommand As String = IMAP_COMMAND_IDENTIFIER + command
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Log(LogTypeEnum.IMAP, sCommand.TrimEnd(IMAP_COMMAND_EOL.ToCharArray()))
            Dim data As Byte() = System.Text.Encoding.ASCII.GetBytes(sCommand.ToCharArray())
            Try
                m_oNetStrm.Write(data, 0, data.Length)
                Dim bRead As Boolean = True
                While bRead
                    Dim sResult As String = m_oRdStrm.ReadLine()
                    sResultArray.Add(sResult)

                    If sResult.StartsWith(IMAP_SERVER_RESPONSE_OK) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
                    ElseIf sResult.StartsWith(IMAP_SERVER_RESPONSE_NO) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    ElseIf sResult.StartsWith(IMAP_SERVER_RESPONSE_BAD) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    Else
                        Log(LogTypeEnum.IMAP, sResult)
                    End If
                End While
            Catch e As Exception
                'LogOut();

                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SERIOUS, e.Message)
            End Try
            Log(LogTypeEnum.IMAP, "")
            Return eImapResponse
        End Function

        ''' <summary>
        '''  retrieve response
        ''' </summary>
        ''' <param name="command">Command to send Imap Server</param>
        ''' <param name="sResultArray">Imap Server response</param>
        ''' <returns>ImapResponseEnum type</returns>
        Protected Function Receive(ByRef sResultArray As ArrayList) As ImapResponseEnum
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Try
                Dim bRead As Boolean = True
                While bRead
                    Dim sResult As String = m_oRdStrm.ReadLine()
                    sResultArray.Add(sResult)

                    If sResult.StartsWith(IMAP_SERVER_RESPONSE_OK) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
                    ElseIf sResult.StartsWith(IMAP_SERVER_RESPONSE_NO) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    ElseIf sResult.StartsWith(IMAP_SERVER_RESPONSE_BAD) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    Else
                        Log(LogTypeEnum.IMAP, sResult)
                    End If
                End While
            Catch e As Exception
                'LogOut();

                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SERIOUS, e.Message)
            End Try
            Log(LogTypeEnum.IMAP, "")
            Return eImapResponse
        End Function

        ''' <summary>
        ''' Send command to server and retrieve response
        ''' </summary>
        ''' <param name="command">Command to send Imap Server</param>
        ''' <param name="sResultArray">Imap Server response</param>
        ''' <returns>ImapResponseEnum type</returns>
        Protected Function SendAndReceiveByNumLines(command As String, ByRef sResultArray As ArrayList, nNumLines As Integer) As ImapResponseEnum
            IMAP_COMMAND_VAL += 1
            Dim nLineCount As Integer = 0
            Dim sCommand As String = IMAP_COMMAND_IDENTIFIER + command
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Log(LogTypeEnum.IMAP, sCommand.TrimEnd(IMAP_COMMAND_EOL.ToCharArray()))
            Dim data As Byte() = System.Text.Encoding.ASCII.GetBytes(sCommand.ToCharArray())
            Try
                m_oNetStrm.Write(data, 0, data.Length)
                Dim bRead As Boolean = True
                While bRead
                    Dim sResult As String = m_oRdStrm.ReadLine()
                    sResultArray.Add(sResult)
                    nLineCount += 1


                    If sResult.StartsWith(IMAP_SERVER_RESPONSE_OK) OrElse nLineCount = nNumLines Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
                    ElseIf sResult.StartsWith(IMAP_SERVER_RESPONSE_NO) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    ElseIf sResult.StartsWith(IMAP_SERVER_RESPONSE_BAD) Then
                        bRead = False
                        Log(LogTypeEnum.IMAP, sResult)
                        eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE
                    Else
                        Log(LogTypeEnum.IMAP, sResult)
                    End If
                End While
            Catch e As Exception
                'LogOut();

                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SERIOUS, e.Message)
            End Try
            Log(LogTypeEnum.IMAP, "")
            Return eImapResponse
        End Function
        ''' <summary>
        ''' Read the Server Response by specified size
        ''' </summary>
        ''' <param name="sBuffer"></param>
        ''' <param name="nSize"></param>
        ''' <returns></returns>
        Protected Function ReceiveBuffer(ByRef sBuffer As String, nSize As Integer) As Integer
            Dim nRead As Integer = -1
            Dim cBuff As Char() = New [Char](nSize + 1) {}
            'cBuff[nSize+1] = '\0';
            nRead = m_oRdStrm.Read(cBuff, 0, nSize)
            Dim sTmp As String = New [String](cBuff)
            sBuffer = sTmp
            Log(LogTypeEnum.IMAP, sBuffer)
            Return nRead
        End Function



        ''' <summary>
        ''' IMAP Capability command
        ''' </summary>
        Public Sub Capability()
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sResultArray As New ArrayList()
            Dim capabilityCommand As String = IMAP_CAPABILITY_COMMAND
            capabilityCommand += IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(capabilityCommand, sResultArray)
                If eImapResponse <> ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_CAPABILITY)
                End If
            Catch e As Exception
                Throw e
            End Try
        End Sub
        ''' <summary>
        ''' Validate Certificate
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="certificate"></param>
        ''' <param name="chain"></param>
        ''' <param name="sslPolicyErrors"></param>
        ''' <returns></returns>
        Public Function ValidateServerCertificate(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
            If sslPolicyErrors = sslPolicyErrors.None Then
                Return True
            End If

            Console.WriteLine("Failed to validate Server Certificate. Error: {0}", sslPolicyErrors)

            Return False
        End Function



    End Class
End Namespace