﻿Imports System.Net.Mail
Imports System.Xml
Imports System.IO
Imports System.Collections
Imports System.Collections.Generic
Imports System

Namespace ImapVB
    ''' <summary>
    ''' Imap class implementes IMAP client API
    ''' </summary>
    Public Class Imap
        Inherits ImapBase
        ''' <summary>
        ''' If user has logged in to his mailbox.
        ''' </summary>
        Private m_bIsLoggedIn As Boolean = False
        ''' <summary>
        ''' Mailbox (Folder) name. Default INBOX.
        ''' </summary>
        Private m_sMailboxName As String = "INBOX"
        ''' <summary>
        ''' If folder is selected.
        ''' </summary>
        Private m_bIsFolderSelected As Boolean = False
        ''' <summary>
        ''' if folder is examined.
        ''' </summary>
        Private m_bIsFolderExamined As Boolean = False
        ''' <summary>
        ''' Total number of messages in mailbox.
        ''' </summary>
        Private m_nTotalMessages As Integer = 0
        ''' <summary>
        ''' Number of recent messages in mailbox.
        ''' </summary>
        Private m_nRecentMessages As Integer = 0
        ''' <summary>
        ''' First unseen message UID
        ''' </summary>
        Private m_nFirstUnSeenMsgUID As Integer = -1


        ''' <summary>
        '''  Login to specified Imap host and default port (143)
        ''' </summary>
        ''' <param name="sHost">Imap Server name</param>
        ''' <param name="sUserId">User's login id</param>
        ''' <param name="sPassword">User's password</param>
        Public Sub Login(sHost As String, sUserId As String, sPassword As String)
            Try
                Login(sHost, IMAP_DEFAULT_PORT, sUserId, sPassword)
            Catch e As Exception
                Throw e
            End Try
        End Sub

        ''' <summary>
        ''' Login to specified Imap host and port
        ''' </summary>
        ''' <param name="sHost">Imap server name</param>
        ''' <param name="nPort">Imap server port</param>
        ''' <param name="sUserId">User's login id</param>
        ''' <param name="sPassword">User's password</param>
        ''' <param name="sslEnabled"> </param>
        ''' <exception cref="IMAP_ERR_LOGIN"
        ''' <exception cref="IMAP_ERR_INVALIDPARAM"
        Public Function Login(sHost As String, nPort As UShort, sUserId As String, sPassword As String, Optional sslEnabled As Boolean = False) As [Boolean]
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim e_login As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_LOGIN, m_sUserId)
            Dim e_invalidparam As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)
            Try

                If sHost.Length = 0 Then
                    Log(LogTypeEnum.[ERROR], "Invalid m_sHost name")
                    Throw e_invalidparam
                End If

                If sUserId.Length = 0 Then
                    Log(LogTypeEnum.[ERROR], "Invalid m_sUserId")
                    Throw e_invalidparam
                End If

                If sPassword.Length = 0 Then
                    Log(LogTypeEnum.[ERROR], "Invalid Password")
                    Throw e_invalidparam
                End If
                If m_bIsConnected Then
                    If m_bIsLoggedIn Then
                        If m_sHost = sHost AndAlso m_nPort = nPort Then
                            If m_sUserId = sUserId AndAlso m_sPassword = sPassword Then
                                Log(LogTypeEnum.INFO, "Connected and Logged in already")
                                Return True
                            Else
                                LogOut()
                            End If
                        Else
                            Disconnect()
                        End If
                    End If
                End If

                m_bIsConnected = False
                m_bIsLoggedIn = False

                Try
                    eImapResponse = Connect(sHost, nPort, sslEnabled)
                    If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                        m_bIsConnected = True
                    Else
                        Return True
                    End If
                Catch e As Exception
                    Throw e
                End Try

                Dim asResultArray As New ArrayList()
                Dim sCommand As String = IMAP_LOGIN_COMMAND
                sCommand += " " + sUserId + " " + sPassword
                sCommand += IMAP_COMMAND_EOL
                Try
                    eImapResponse = SendAndReceive(sCommand, asResultArray)
                    If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                        m_bIsLoggedIn = True
                        m_sUserId = sUserId
                        m_sPassword = sPassword
                    Else
                        Throw e_login

                    End If
                Catch e As Exception
                    Throw e
                End Try
                Return True
            Catch ex As Exception
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Logout the user: It logout the user and disconnect the connetion from IMAP server.
        ''' </summary>
        Public Sub LogOut()
            If m_bIsLoggedIn Then
                Dim eImapResponse As ImapResponseEnum
                Dim asResultArray As New ArrayList()
                Dim sCommand As String = IMAP_LOGOUT_COMMAND
                sCommand += IMAP_COMMAND_EOL
                Try
                    eImapResponse = SendAndReceive(sCommand, asResultArray)
                Catch e As Exception
                    Disconnect()
                    m_bIsLoggedIn = False
                    Throw e
                End Try
                Disconnect()
                m_bIsLoggedIn = False
            End If
        End Sub

        ''' <summary>
        ''' Select the sFolder/mailbox after login
        ''' </summary>
        ''' <param name="sFolder">mailbox folder</param>
        ''' <exception cref="IMAP_ERR_SELECT"
        ''' <exception cref="IMAP_ERR_INSUFFICIENT_DATA"
        ''' <exception cref="IMAP_ERR_INVALIDPARAM"
        Public Sub SelectFolder(sFolder As String)
            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED, e.Message)

                End Try
            End If
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim e_select As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SELECT, sFolder)
            Dim e_invalidparam As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)
            If sFolder.Length = 0 Then
                Throw e_invalidparam
            End If
            'If m_bIsFolderSelected Then
            '    If m_sMailboxName = sFolder Then
            '        Log(LogTypeEnum.INFO, "Folder is already selected")
            '        Return
            '    Else
            '        m_bIsFolderSelected = False
            '    End If
            'End If
            Dim asResultArray As New ArrayList()
            Dim sCommand As String = IMAP_SELECT_COMMAND
            sCommand += " " + sFolder + IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    m_sMailboxName = sFolder
                    m_bIsFolderSelected = True
                Else
                    Throw e_select
                End If
            Catch e As Exception
                Throw e
            End Try

            '-------------------------
            ' PARSE RESPONSE

            Dim bResult As Boolean = False
            For Each sLine As String In asResultArray
                ' If this is an unsolicited response starting with '*'
                If sLine.IndexOf(IMAP_UNTAGGED_RESPONSE_PREFIX) <> -1 Then
                    ' parse the line by space
                    Dim asTokens As String()
                    asTokens = sLine.Split(" "c)
                    If asTokens(2) = "EXISTS" Then
                        ' The line will look like "* 2 EXISTS"
                        m_nTotalMessages = Convert.ToInt32(asTokens(1))
                    ElseIf asTokens(2) = "RECENT" Then
                        ' The line will look like "* 2 RECENT"
                        m_nRecentMessages = Convert.ToInt32(asTokens(1))
                    ElseIf asTokens(2) = "[UNSEEN" Then
                        ' The line will look like "* OK [UNSEEN 2]"
                        Dim sUIDPart As String = asTokens(3).Substring(0, asTokens(3).Length - 1)
                        m_nFirstUnSeenMsgUID = Convert.ToInt32(sUIDPart)
                    End If
                    ' If this line looks like "<command-tag> OK ..."
                ElseIf sLine.IndexOf(IMAP_SERVER_RESPONSE_OK) <> -1 Then
                    bResult = True
                    Exit For
                End If
            Next

            If Not bResult Then
                Throw e_select
            End If

            Dim sLogStr As String = "TotalMessages[" + m_nTotalMessages.ToString() + "] ,"
            sLogStr += "RecentMessages[" + m_nRecentMessages.ToString() + "] ,"
            If m_nFirstUnSeenMsgUID > 0 Then
                sLogStr += "FirstUnSeenMsgUID[" + m_nFirstUnSeenMsgUID.ToString() + "] ,"
            End If
            'Log(LogTypeEnum.INFO, sLogStr);

        End Sub

        ''' <summary>
        ''' Examine the sFolder/mailbox after login
        ''' </summary>
        ''' <param name="sFolder">Mailbox folder</param>
        Public Sub ExamineFolder(sFolder As String)
            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim e_examine As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_EXAMINE, sFolder)
            Dim e_invalidparam As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)
            If sFolder.Length = 0 Then
                Throw e_invalidparam
            End If
            If m_bIsFolderExamined Then
                If m_sMailboxName = sFolder Then
                    Log(LogTypeEnum.INFO, "Folder is already selected")
                    Return
                Else
                    m_bIsFolderExamined = False
                End If
            End If
            Dim asResultArray As New ArrayList()
            Dim sCommand As String = IMAP_EXAMINE_COMMAND
            sCommand += " " + sFolder + IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    m_sMailboxName = sFolder
                    m_bIsFolderExamined = True
                Else
                    Throw e_examine
                End If
            Catch e As Exception
                Throw e
            End Try
            '-------------------------
            ' PARSE RESPONSE

            Dim bResult As Boolean = False
            For Each sLine As String In asResultArray
                ' If this is an unsolicited response starting with '*'
                If sLine.IndexOf(IMAP_UNTAGGED_RESPONSE_PREFIX) <> -1 Then
                    ' parse the line by space
                    Dim asTokens As String()
                    asTokens = sLine.Split(" "c)
                    If asTokens(2) = "EXISTS" Then
                        ' The line will look like "* 2 EXISTS"
                        m_nTotalMessages = Convert.ToInt32(asTokens(1))
                    ElseIf asTokens(2) = "RECENT" Then
                        ' The line will look like "* 2 RECENT"
                        m_nRecentMessages = Convert.ToInt32(asTokens(1))
                    ElseIf asTokens(2) = "[UNSEEN" Then
                        ' The line will look like "* OK [UNSEEN 2]"
                        Dim sUIDPart As String = asTokens(3).Substring(0, asTokens(3).Length - 1)
                        m_nFirstUnSeenMsgUID = Convert.ToInt32(sUIDPart)
                    End If
                    ' If this line looks like "<command-tag> OK ..."
                ElseIf sLine.IndexOf(IMAP_SERVER_RESPONSE_OK) <> -1 Then
                    bResult = True
                    Exit For
                End If
            Next

            If Not bResult Then
                Throw e_examine
            End If

            Dim sLogStr As String = "TotalMessages[" + m_nTotalMessages.ToString() + "] ,"
            sLogStr += "RecentMessages[" + m_nRecentMessages.ToString() + "] ,"
            If m_nFirstUnSeenMsgUID > 0 Then
                sLogStr += "FirstUnSeenMsgUID[" + m_nFirstUnSeenMsgUID.ToString() + "] ,"
            End If
            'Log(LogTypeEnum.INFO, sLogStr);

        End Sub

        ''' <summary>
        ''' Restore the connection using available old data
        ''' Select the sFolder if previously selected
        ''' </summary>
        ''' <param name="bSelectFolder">If true then it selects the folder</param>
        Sub Restore(bSelectFolder As Boolean)
            Dim e_insufficiantdata As New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA)
            If m_sHost.Length = 0 OrElse m_sUserId.Length = 0 OrElse m_sPassword.Length = 0 Then
                Throw e_insufficiantdata
            End If
            Try
                m_bIsLoggedIn = False
                Login(m_sHost, m_nPort, m_sUserId, m_sPassword)
                If bSelectFolder AndAlso m_sMailboxName.Length > 0 Then
                    If m_bIsFolderSelected Then
                        m_bIsFolderSelected = False
                        SelectFolder(m_sMailboxName)
                    ElseIf m_bIsFolderExamined Then
                        m_bIsFolderExamined = False
                        ExamineFolder(m_sMailboxName)
                    Else
                        SelectFolder(m_sMailboxName)
                    End If
                End If
            Catch e As Exception
                Throw e
            End Try

        End Sub

        ''' <summary>
        ''' Check if enough quota is available
        ''' </summary>
        ''' <param name="sFolderName">Mailbox folder</param>
        ''' <returns>true if enough mail quota</returns>
        Public Function HasEnoughQuota(sFolderName As String) As Boolean
            Try
                Dim bUnlimitedQuota As Boolean = False


                Dim nUsedKBytes As Integer = 0
                Dim nTotalKBytes As Integer = 0

                GetQuota(sFolderName, bUnlimitedQuota, nUsedKBytes, nTotalKBytes)

                If bUnlimitedQuota OrElse (nUsedKBytes < nTotalKBytes) Then
                    Return True
                Else
                    Return False
                End If
            Catch e As ImapException
                Throw e
            End Try
        End Function

        ''' <summary>
        ''' Get the quota for specific folder
        ''' </summary>
        ''' <param name="sFolderName">Mailbox folder</param>
        ''' <param name="bUnlimitedQuota">Is unlimited quota</param>
        ''' <param name="nUsedKBytes">Used quota in Kbytes</param>
        ''' <param name="nTotalKBytes">Total quota in KBytes</param>
        Public Sub GetQuota(sFolderName As String, ByRef bUnlimitedQuota As Boolean, ByRef nUsedKBytes As Integer, ByRef nTotalKBytes As Integer)
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim bResult As Boolean = False
            bUnlimitedQuota = False
            nUsedKBytes = 0
            nTotalKBytes = 0
            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            If sFolderName.Length = 0 Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)
            End If

            Dim asResultArray As New ArrayList()
            Dim sCommand As String = IMAP_GETQUOTA_COMMAND
            sCommand += " " + sFolderName + IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    m_sMailboxName = sFolderName
                    m_bIsFolderExamined = True
                    Dim quotaPrefix As String = IMAP_UNTAGGED_RESPONSE_PREFIX + " "
                    quotaPrefix += IMAP_QUOTA_RESPONSE + " "
                    For Each sLine As String In asResultArray
                        If sLine.StartsWith(quotaPrefix) = True Then
                            ' Find the open and close paranthesis, and extract
                            ' the part inside out.
                            Dim nStart As Integer = sLine.IndexOf("("c)
                            Dim nEnd As Integer = sLine.IndexOf(")"c, nStart)
                            If nStart <> -1 AndAlso nEnd <> -1 AndAlso nEnd > nStart Then
                                Dim sQuota As String = sLine.Substring(nStart + 1, nEnd - nStart - 1)
                                If sQuota.Length > 0 Then
                                    ' Parse the space-delimited quota information which
                                    ' will look like "STORAGE <used> <total>"
                                    Dim asArrList As String()
                                    ' = new ArrayList();
                                    asArrList = sQuota.Split(" "c)

                                    ' get the used and total kbytes from these tokens
                                    If asArrList.Length = 3 AndAlso asArrList(0) = "STORAGE" Then
                                        nUsedKBytes = Convert.ToInt32(asArrList(1), 10)


                                        nTotalKBytes = Convert.ToInt32(asArrList(2), 10)
                                    Else
                                        Dim [error] As String = "Invalid Quota information :" + sQuota
                                        Log(LogTypeEnum.[ERROR], [error])
                                        Exit For
                                    End If
                                Else
                                    bUnlimitedQuota = True
                                End If
                            Else
                                Dim [error] As String = "Invalid Quota IMAP Response : " + sLine
                                Log(LogTypeEnum.[ERROR], [error])
                                Exit For
                            End If
                            ' If the line looks like "<command-tag> OK ..."
                        ElseIf sLine.IndexOf(IMAP_SERVER_RESPONSE_OK) <> -1 Then
                            bResult = True
                            Exit For
                        End If
                    Next

                    If Not bResult Then
                        Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_QUOTA)
                    End If
                    If bUnlimitedQuota Then
                        Log(LogTypeEnum.INFO, "GETQUOTA quota=[unlimited].")
                    Else
                        'Log(LogTypeEnum.INFO, sLogStr);
                        Dim sLogStr As String = "GETQUOTA used=[" + nUsedKBytes.ToString() + "], total=[" + nTotalKBytes.ToString() + "]"
                    End If
                End If
            Catch e As Exception
                Throw e
            End Try

        End Sub

        ''' <summary>
        ''' Store flag
        ''' </summary>
        ''' <param name="sUid"></param>
        ''' <param name="flag"> E.g \Deleted </param>
        ''' <param name="removeFlag">Remove the flaf </param>
        Public Sub SetFlag(sUid As String, flag As String, Optional removeFlag As Boolean = False)
            If [String].IsNullOrEmpty(sUid) Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM, "Invalid uid")
            End If
            If [String].IsNullOrEmpty(flag) Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM, "Invalid flag")
            End If
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE


            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            If Not m_bIsFolderSelected AndAlso Not m_bIsFolderExamined Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED)
            End If

            Dim asResultArray As New ArrayList()
            '            string sCommand = IMAP_UIDSTORE_COMMAND;
            Dim sCommand As String = IMAP_STORE_COMMAND
            sCommand += " " + sUid + " " + (If(removeFlag, IMAP_REMOVEFLAGS_COMMAND, IMAP_SETFLAGS_COMMAND)) + " " + flag + IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse <> ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SEARCH, asResultArray(0).ToString())
                End If
            Catch e As Exception
                Throw e
            End Try
        End Sub

        ''' <summary>
        ''' Expunge
        ''' </summary>
        Public Sub Expunge()

            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE


            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            If Not m_bIsFolderSelected AndAlso Not m_bIsFolderExamined Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED)
            End If

            Dim asResultArray As New ArrayList()
            Dim sCommand As String = IMAP_EXPUNGE_COMMAND
            sCommand += IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse <> ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SEARCH, asResultArray(0).ToString())
                End If
            Catch e As Exception
                Throw e
            End Try
        End Sub

        ''' <summary>
        ''' Move message to specified folder
        ''' </summary>
        ''' <param name="sUid">UID of the message</param>
        ''' <param name="sFolderName"> Folder where you want to move the message</param>
        Public Sub MoveMessage(sUid As String, sFolderName As String)
            CopyMessage(sUid, sFolderName)
            Expunge()
            SetFlag(sUid, "\Answered")
            Expunge()
        End Sub
        Public Sub DeleteMessage2(sUid As String)
            SetFlag(sUid, "\Deleted")
            Expunge()
        End Sub
        Public Sub ValidateMessage(sUid As String)
            SetFlag(sUid, "\Seen")
            Expunge()
        End Sub



        ''' <summary>
        ''' Copy Message
        ''' </summary>
        ''' <param name="sUid">Either UID or range of uid e.g 1:2</param>
        ''' <param name="sFolderName">Folder where it needs to be copied</param>
        Public Sub CopyMessage(sUid As String, sFolderName As String)
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE

            If [String].IsNullOrEmpty(sFolderName) Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM, "Invalid folder name.")
            End If
            If [String].IsNullOrEmpty(sUid) Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM, "Invalid uid")
            End If
            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If

            If Not m_bIsFolderSelected AndAlso Not m_bIsFolderExamined Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED)
            End If

            Dim asResultArray As New ArrayList()
            Dim sCommand As String = IMAP_COPY_COMMAND
            sCommand += " " + sUid + " " + sFolderName + IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse <> ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SEARCH, asResultArray(0).ToString())
                End If
            Catch e As Exception
                Throw e
            End Try

        End Sub

        ''' <summary>
        ''' Get the message size
        ''' </summary>
        ''' <param name="sUid"></param>
        ''' <returns>message size</returns>
        Public Function GetMessageSize(sUid As String) As Long
            If [String].IsNullOrEmpty(sUid) Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM, "Invalid uid")
            End If
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE


            If Not m_bIsLoggedIn Then
                Try
                    Restore(False)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            If Not m_bIsFolderSelected AndAlso Not m_bIsFolderExamined Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED)
            End If

            Dim asResultArray As New ArrayList()
            Dim sCommand As String = IMAP_FETCH_COMMAND
            sCommand += " " + sUid + " " + IMAP_RFC822_SIZE_COMMAND + IMAP_COMMAND_EOL
            Try
                eImapResponse = SendAndReceive(sCommand, asResultArray)
                If eImapResponse <> ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SEARCH, asResultArray(0).ToString())
                End If
                Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                Dim sBodyStruct As String = ""
                Dim bResult As Boolean = False
                Dim nStart As Integer = -1
                For Each sLine As String In asResultArray
                    nStart = sLine.IndexOf(IMAP_FETCH_COMMAND)
                    If sLine.StartsWith(IMAP_UNTAGGED_RESPONSE_PREFIX) AndAlso (nStart <> -1) Then
                        sBodyStruct = sLine
                    ElseIf sLine.StartsWith(sLastLine) Then
                        bResult = True
                        Exit For
                    End If
                Next
                If Not bResult Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHSIZE)
                End If
                nStart = sBodyStruct.IndexOf(IMAP_RFC822_SIZE_COMMAND)
                Dim nEnd As Integer = sBodyStruct.IndexOf(")")
                Dim size As String = sBodyStruct.Substring(nStart + IMAP_RFC822_SIZE_COMMAND.Length, nEnd - (nStart + IMAP_RFC822_SIZE_COMMAND.Length))


                Return Convert.ToUInt32(size.Trim())
            Catch e As Exception
                Throw e
            End Try
        End Function



        ''' <summary>
        ''' Search the messages by specified criterias
        ''' </summary>
        ''' <param name="asSearchData">Search criterias</param>
        ''' <param name="bExactMatch">Is it exact search</param>
        ''' <param name="asSearchResult">search result</param>
        Public Sub SearchMessage(asSearchData As String(), bExactMatch As Boolean, asSearchResult As ArrayList)
            If Not m_bIsLoggedIn Then
                Try
                    Restore(True)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            If Not m_bIsFolderSelected AndAlso Not m_bIsFolderExamined Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED)
            End If
            Dim nCount As Integer = asSearchData.Length
            If nCount = 0 Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)
            End If
            '--------------------------
            ' PREPARE SEARCH KEY/VALUE

            Dim sCommandSuffix As String = ""
            For Each sLine As String In asSearchData
                If sLine.Length = 0 Then
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)


                End If

                ' convert to lower case once for all
                sLine.ToLower()

                If sCommandSuffix.Length > 0 Then
                    sCommandSuffix += " "
                End If
                sCommandSuffix += sLine
            Next

            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandString As String = IMAP_SEARCH_COMMAND + " " + sCommandSuffix
            sCommandString += IMAP_COMMAND_EOL
            Dim asResultArray As New ArrayList()
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    '-------------------------
                    ' PARSE RESPONSE
                    nCount = asResultArray.Count
                    Dim bResult As Boolean = False
                    Dim sPrefix As String = IMAP_UNTAGGED_RESPONSE_PREFIX + " "
                    sPrefix += IMAP_SEARCH_RESPONSE
                    For Each sLine As String In asResultArray
                        Dim nPos As Integer = sLine.IndexOf(sPrefix)
                        If nPos <> -1 Then
                            nPos += sPrefix.Length
                            Dim sSuffix As String = sLine.Substring(nPos)
                            sSuffix.Trim()
                            Dim asSearchRes As String() = sSuffix.Split(" "c)
                            For Each sResultLine As String In asSearchRes
                                asSearchResult.Add(sResultLine)
                            Next
                            bResult = True
                            Exit For
                        End If
                    Next
                    If Not bResult Then
                        Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SEARCH, sCommandSuffix)
                    End If
                Else
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_SEARCH, asResultArray(0).ToString())
                End If
            Catch e As ImapException
                LogOut()
                Throw e
            End Try

        End Sub
        ''' <summary>
        ''' Fetch the body for specified part
        ''' </summary>
        ''' <param name="sMessageUID"> Message uid</param>
        ''' <param name="sMessagePart">Message part</param>
        ''' <param name="aArray">Body data as Array</param>
        'public void FetchPartBody(string sMessageUID,
        '    string sMessagePart,
        '    ref string sData)
        '{
        '    if (!m_bIsLoggedIn)
        '    {
        '        try
        '        {
        '            Restore(true);
        '        }
        '        catch (ImapException e)
        '        {
        '            if (e.Type != ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA)
        '                throw e;

        '            throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED);
        '        }
        '    }
        '    if (!m_bIsFolderSelected && !m_bIsFolderExamined) 
        '    {
        '        throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED);
        '    }
        '    if (sMessageUID.Length == 0)
        '        throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM);
        '    string sPart = sMessagePart;
        '    if (sPart.Length == 0 )
        '        sPart = IMAP_MSG_DEFAULT_PART;
        '    try
        '    {

        '        GetBody(sMessageUID,sPart, ref sData);

        '    }
        '    catch (ImapException e)
        '    {
        '        throw e;
        '    }
        '}
        ''' <summary>
        ''' Fetch Header of the message uid and part
        ''' </summary>
        ''' <param name="sMessageUID"> Message UID</param>
        ''' <param name="sMessagePart"> Message part</param>
        ''' <param name="asMessageHeader">Output Array</param>
        Public Sub FetchPartHeader(sMessageUID As String, sMessagePart As String, asMessageHeader As ArrayList)
            If Not m_bIsLoggedIn Then
                Try
                    Restore(True)
                Catch e As ImapException
                    If e.Type <> ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA Then
                        Throw e
                    End If

                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED)
                End Try
            End If
            If Not m_bIsFolderSelected AndAlso Not m_bIsFolderExamined Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED)
            End If
            If sMessageUID.Length = 0 Then
                Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDPARAM)
            End If
            Dim sPart As String = sMessagePart
            If sPart.Length = 0 Then
                sPart = IMAP_MSG_DEFAULT_PART
            End If
            Try
                GetHeader(sMessageUID, sPart, asMessageHeader)
            Catch e As ImapException
                Throw e
            End Try
        End Sub

        ''' <summary>
        ''' Fetch the full message
        ''' </summary>
        ''' <param name="sMessageUID">Message UID </param>
        ''' <param name="oXmlDoc">Message is stored as XmlDocument object</param>
        'public void FetchMessage(string sMessageUID, XmlTextWriter oXmlWriter, bool bFetchBody)
        '{
        '    if (!m_bIsLoggedIn)
        '    {
        '        try
        '        {
        '            Restore(true);
        '        }
        '        catch (ImapException e)
        '        {
        '            if (e.Type != ImapException.ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA)
        '                throw e;

        '            throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTCONNECTED);
        '        }
        '    }
        '    if (!m_bIsFolderSelected && !m_bIsFolderExamined) 
        '    {
        '        throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_NOTSELECTED);
        '    }
        '    try 
        '    {
        '        String sheader;
        '        sheader = GetSubject( sMessageUID);
        '        oXmlWriter.WriteStartElement("Part");
        '        oXmlWriter.WriteAttributeString("ID", "0");
        '        oXmlWriter.WriteElementString("Subject",sheader);

        '            oXmlWriter.WriteStartElement("Part");
        '            oXmlWriter.WriteAttributeString("ID", "1");

        '        GetBodyStructure(sMessageUID, oXmlWriter,bFetchBody);
        '        oXmlWriter.WriteEndElement();
        '    }
        '    catch (ImapException e)
        '    {
        '        throw e;
        '    }
        '}

        ''' <summary>
        ''' Get the Body structure of the message.
        ''' If message is single part then first part is 1
        ''' If message is multipart then first part is 0
        ''' </summary>
        ''' <param name="sMessageUID"> Message UID</param>
        ''' <param name="oXmlElem"> Body structure as XmlElement</param>
        'public void GetBodyStructure(string sMessageUID, XmlTextWriter oXmlWriter, bool bFetchBody)
        '{
        '    ImapResponseEnum eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE;
        '    //string sCommandSuffix = sMessageUID + " " + "BODYSTRUCTURE";
        '    string sCommandSuffix = sMessageUID + " " + "BODY[TEXT]";
        '    string sCommandString = IMAP_FETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL;

        '    try
        '    {
        '        //-----------------------
        '        // SEND SEARCH REQUEST
        '        ArrayList asResultArray = new ArrayList();
        '        eImapResponse = SendAndReceive(sCommandString, ref asResultArray);
        '        if (eImapResponse == ImapResponseEnum.IMAP_SUCCESS_RESPONSE)
        '        {
        '            string sLastLine = IMAP_SERVER_RESPONSE_OK;
        '            string sBodyStruct = "";
        '            bool bResult = false;
        '            int nStart = -1;
        '            foreach (string sLine in asResultArray)
        '            {
        '                nStart = sLine.IndexOf(IMAP_BODYSTRUCTURE_COMMAND);
        '                if (sLine.StartsWith(IMAP_UNTAGGED_RESPONSE_PREFIX) &&
        '                    (nStart != -1))
        '                {
        '                    sBodyStruct = sLine;
        '                }
        '                else if (sLine.StartsWith(sLastLine))
        '                {
        '                    bResult = true;
        '                    break;
        '                }
        '            }
        '            if (!bResult)
        '            {
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHBODYSTRUCT);
        '            }
        '            else if (sBodyStruct.Length == 0)
        '            {
        '                Log(LogTypeEnum.INFO, "Bodystructure is empty");
        '                return;
        '            }
        '            nStart = sBodyStruct.IndexOf(IMAP_BODYSTRUCTURE_COMMAND);
        '            sBodyStruct = sBodyStruct.Substring(nStart + IMAP_BODYSTRUCTURE_COMMAND.Length);
        '            int nEnd = sBodyStruct.LastIndexOf(")");
        '            if (nEnd == -1)
        '            {
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHBODYSTRUCT);
        '            }
        '            sBodyStruct = sBodyStruct.Substring(0, nEnd);
        '            string sPartPrefix = "";
        '            if (!ParseBodyStructure(sMessageUID, ref sBodyStruct, oXmlWriter, sPartPrefix, bFetchBody))
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHBODYSTRUCT);
        '        }
        '        else
        '        {
        '            throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG, sCommandSuffix);
        '        }

        '    }
        '    catch (ImapException e)
        '    {
        '        LogOut();
        '        throw e;
        '    }
        '}

        ''' <summary>
        ''' Parse the bodystructure and store as XML Element
        ''' </summary>
        ''' <param name="sBodyStruct">Bosy Structure</param>
        ''' <param name="oXmlBodyPart">Body Structure in XML</param>
        ''' <param name="sPartPrefix">Part Prefix</param>
        ''' <returns>true/false</returns>
        'bool ParseBodyStructure(string sMessageUID, ref string sBodyStruct, XmlTextWriter oXmlBodyPart,
        '    string sPartPrefix, bool bFetchBody)
        '{

        '    bool bMultiPart = false;
        '    string sTemp = "";
        '    ArrayList asAttrs = new ArrayList();
        '    sBodyStruct = sBodyStruct.Trim();

        '    //Check if this is NIL
        '    if (IsNilString (ref sBodyStruct))
        '        return true;
        '    //Look for '('
        '    if (sBodyStruct[0] == '(')
        '    {
        '        //Check if multipart or singlepart.
        '        //Multipart will have another '(' here
        '        // and single part will not.
        '        char ch;
        '        ch = sBodyStruct[1];
        '        if (ch != '(')
        '        {
        '            //Singal part
        '            if (ch == ')')
        '            {
        '                sBodyStruct = sBodyStruct.Substring(2);
        '                return true;
        '            }
        '            //remove opening paranthesis
        '            sBodyStruct = sBodyStruct.Substring(1);
        '            string sType = "";
        '            string sSubType = "";
        '            if (!GetContentType(ref sBodyStruct, ref sType, ref sSubType, ref sTemp))
        '            {
        '                return false;
        '            }
        '            asAttrs.Add(IMAP_MESSAGE_CONTENT_TYPE);
        '            asAttrs.Add(sTemp);

        '            // Message-id (optional)
        '            if (!ParseQuotedString( ref sBodyStruct, ref sTemp )) 
        '            {
        '                Log(LogTypeEnum.ERROR, "Failed getting Message Id.");
        '                return false;
        '            }
        '            if (sTemp.Length > 0) 
        '            {
        '                asAttrs.Add(IMAP_MESSAGE_ID);
        '                asAttrs.Add(sTemp);
        '            }
        '            // Content-Description (optional)
        '            if (!ParseQuotedString( ref sBodyStruct, ref sTemp )) 
        '            {
        '                Log(LogTypeEnum.ERROR, "Failed getting the Content Desc.");
        '                return false;
        '            }
        '            if (sTemp.Length > 0) 
        '            {
        '                asAttrs.Add(IMAP_MESSAGE_CONTENT_DESC);
        '                asAttrs.Add(sTemp);
        '            }

        '            // Content-Transfer-Encoding
        '            if (!ParseQuotedString( ref sBodyStruct, ref sTemp )) 
        '            {
        '                Log(LogTypeEnum.ERROR, "Failed getting the Content Encoding.");
        '                return false;
        '            }
        '            asAttrs.Add(IMAP_MESSAGE_CONTENT_ENCODING);
        '            asAttrs.Add(sTemp);

        '            // Content Size in bytes
        '            if (!ParseString( ref sBodyStruct, ref sTemp )) 
        '            {
        '                Log(LogTypeEnum.ERROR, "Failed getting the Content Size.");
        '                return false;
        '            }
        '            asAttrs.Add(IMAP_MESSAGE_CONTENT_SIZE);
        '            asAttrs.Add(sTemp);
        '            sTemp = sType + "/" + sSubType;
        '            if (sTemp.ToLower() == IMAP_MESSAGE_RFC822.ToLower()) //email attachment
        '            {
        '                if (!ParseEnvelope(ref sBodyStruct, asAttrs )) 
        '                {
        '                    Log(LogTypeEnum.ERROR, "Failed getting the Message Envelope.");
        '                    return false;
        '                }

        '                if (!ParseBodyStructure(sMessageUID, ref sBodyStruct, oXmlBodyPart,
        '                    sPartPrefix, bFetchBody )) 
        '                {
        '                    Log (LogTypeEnum.ERROR, "Failed getting Attached Message.");
        '                    return false;
        '                }
        '                // No of Lines in the message
        '                if (!ParseString(ref sBodyStruct, ref sTemp )) 
        '                {
        '                    Log(LogTypeEnum.ERROR, "Failed getting the Content Lines.");
        '                    return false;
        '                }
        '                if (sTemp.Length > 0) 
        '                {
        '                    asAttrs.Add(IMAP_MESSAGE_CONTENT_LINES);
        '                    asAttrs.Add(sTemp);
        '                }
        '            }
        '            else if (sType == "text") //simple text
        '            {
        '                // No of Lines in the message
        '                if (!ParseString(ref sBodyStruct, ref sTemp )) 
        '                {
        '                    Log(LogTypeEnum.ERROR, "Failed getting the Content Lines.");
        '                    return false;
        '                }
        '                if (sTemp.Length > 0) 
        '                {
        '                    asAttrs.Add(IMAP_MESSAGE_CONTENT_LINES);
        '                    asAttrs.Add(sTemp);
        '                }
        '            }
        '            // MD5 CRC Error Check
        '            // Don't know what to do with it
        '            if (sBodyStruct[0] == ' ') 
        '            {
        '                if (!ParseString( ref sBodyStruct, ref sTemp ))
        '                   return false;
        '            }
        '        }
        '        else // MULTIPART
        '        {
        '            bMultiPart = true;
        '            // remove the open paranthesis
        '            sBodyStruct = sBodyStruct.Substring(1);
        '            uint nPartNumber = 0;
        '            string sPartNumber= "";
        '            do 
        '            {
        '                // prepare next part number
        '                nPartNumber++;

        '                if (sPartPrefix.Length > 0) 
        '                    sPartNumber = sPartPrefix + "." +  nPartNumber.ToString();
        '                else
        '                    sPartNumber = nPartNumber.ToString();

        '                //XmlElement oXmlChildPart;
        '                oXmlBodyPart.WriteStartElement("Part");
        '                oXmlBodyPart.WriteAttributeString("ID",sPartNumber);
        '                // add a new child to the part with
        '                // an empty attribute array. This array will be filled
        '                // in the "if" condition block.
        '                int nCount = asAttrs.Count;
        '                for (int i = 0; i < nCount; i = i + 2)
        '                {
        '                    oXmlBodyPart.WriteElementString(asAttrs[i].ToString(),asAttrs[i + 1].ToString());
        '                }
        '                if (sPartNumber.Length > 0 && 
        '                    sPartNumber != "0" &&
        '                    bFetchBody == true)
        '                {
        '                    try
        '                    {
        '                        string sData = "";
        '                        GetBody(sMessageUID, sPartNumber, ref sData);
        '                        if (sData.Length > 0 &&
        '                            ((sData.ToLower()).IndexOf(IMAP_MESSAGE_CONTENT_TYPE.ToLower()) == -1))
        '                        {
        '                            oXmlBodyPart.WriteElementString("DATA", sData);
        '                        }
        '                    }
        '                    catch (ImapException e)
        '                    {
        '                        Log(LogTypeEnum.ERROR, "Exception:Invalid Body Structure. Error:" + e.Message);
        '                        return false;
        '                    }
        '                }


        '                // add a new child to the part with
        '                // an empty attribute array. This array will be filled
        '                // in the "if" condition block.
        '                //oXmlBodyPart.AppendChild(oXmlChildPart);
        '                // parse this body part
        '                if (!ParseBodyStructure(sMessageUID, ref sBodyStruct, oXmlBodyPart,
        '                    sPartNumber, bFetchBody)) 
        '                {
        '                    return false;
        '                }
        '              oXmlBodyPart.WriteEndElement();
        '            }
        '            while (sBodyStruct[0] == '('); // For each body part
        '            // Content-type
        '            string sType = IMAP_MESSAGE_MULTIPART;
        '            string sSubType = "";
        '            if (!GetContentType(ref sBodyStruct, ref sType, ref sSubType,
        '                ref sTemp )) 
        '            {
        '                return false;
        '            }
        '            asAttrs.Add(IMAP_MESSAGE_CONTENT_TYPE);
        '            asAttrs.Add(sTemp);
        '        }
        '        //----------------------------------
        '        // COMMON FOR SINGLE AND MULTI PART

        '        // Disposition
        '        if (sBodyStruct[0] == ' ') 
        '        {
        '            if (!GetContentDisposition(ref sBodyStruct, ref sTemp )) 
        '            {
        '                Log(LogTypeEnum.ERROR, "Failed getting the Content Disp.");
        '                return false;
        '            }
        '            if (sTemp.Length > 0) 
        '            {
        '                asAttrs.Add(IMAP_MESSAGE_CONTENT_DISP);
        '                asAttrs.Add(sTemp);
        '            }
        '        }
        '        // Language
        '        if (sBodyStruct[0] == ' ')
        '        {
        '            if (!ParseLanguage(ref sBodyStruct, ref sTemp ))
        '                return false;
        '        }
        '        // Extension data
        '        while (sBodyStruct[0] == ' ')
        '            if (!ParseExtension(ref sBodyStruct, ref sTemp ))
        '                return false;

        '        // this is the end of the body part
        '        if (!FindAndRemove(ref sBodyStruct, ')' ))
        '            return false;

        '        // Finally, set the attribute array to the part
        '        // EXCEPTION : if multipart and this is the root
        '        // part, the header is already prepared in the
        '        // GetBodyStructure function and hence DO NOT set it.

        '        if (!bMultiPart || sPartPrefix.Length > 0)
        '        {
        '            int nCount = asAttrs.Count;
        '            for (int i = 0; i < nCount ; i= i+2)
        '            {
        '                oXmlBodyPart.WriteElementString(asAttrs[i].ToString(), asAttrs[i+1].ToString());
        '            }

        '        }
        '        return true;
        '    }

        '    Log (LogTypeEnum.ERROR, "Invalid Body Structure");
        '    return false;
        '}
        ''' <summary>
        ''' Get the message body for specified part
        ''' </summary>
        ''' <param name="sMessageUid">Message uid</param>
        ''' <param name="sPartNumber">Body part</param>
        ''' <param name="aArray">Out data</param>
        'void GetBody(string sMessageUid, string sPartNumber, ref string sData)
        '{
        '    ImapResponseEnum eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE;
        '    string sCommandSuffix = "";
        '    sCommandSuffix = sMessageUid + " " + "BODY[" + sPartNumber + "]";
        '    string sCommandString = IMAP_UIDFETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL ;

        '    try
        '    {
        '        //-----------------------
        '        // SEND SEARCH REQUEST
        '        ArrayList asResultArray = new ArrayList();
        '        eImapResponse = SendAndReceiveByNumLines(sCommandString, ref asResultArray, 1);
        '        if (eImapResponse == ImapResponseEnum.IMAP_SUCCESS_RESPONSE)
        '        {
        '            //-------------------------
        '            // PARSE RESPONSE
        '            string sLastLine = IMAP_COMMAND_IDENTIFIER + " " + IMAP_OK_RESPONSE;
        '            string sLine = asResultArray[0].ToString();
        '            if (!sLine.StartsWith(IMAP_UNTAGGED_RESPONSE_PREFIX))
        '            {
        '                string sLog = "InValid Message Response " + sLine + ".";
        '                Log (LogTypeEnum.ERROR, sLog);
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG);
        '            }
        '            long lMessageSize = GetResponseSize(sLine);
        '            if (lMessageSize == 0L)
        '            {
        '                string sLog = "InValid Message Response " + sLine + ".";
        '                Log (LogTypeEnum.ERROR, sLog);
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG);
        '            }
        '            sData = "";
        '            ReceiveBuffer(ref sData,Convert.ToInt32(lMessageSize));
        '            if (sData.Length == 0) 
        '            {
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG, sCommandSuffix ); 
        '            }
        '            //Convert.FromBase64String(sData).ToString();
        '            eImapResponse = Receive(ref asResultArray);
        '            if (eImapResponse != ImapResponseEnum.IMAP_SUCCESS_RESPONSE)
        '                throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG, sCommandSuffix );

        '           /***
        '            int nCount = asResultArray.Count;
        '            for (int i = 1; i < nCount; i++)
        '            {
        '                string sTmpData = asResultArray[i].ToString();
        '                if (sTmpData.Length > 0)
        '                {
        '                    if (sTmpData[0] == ')' ||
        '                        (sTmpData.IndexOf(IMAP_SERVER_RESPONSE_OK) != -1))
        '                    {
        '                        break;
        '                    }
        '                    else
        '                        sData += sTmpData;
        '                }

        '            }****/
        '        }
        '        else
        '            throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG, sCommandSuffix );
        '    }
        '    catch (ImapException e)
        '    {
        '        LogOut();
        '        throw e;
        '    }	
        '}

        ''' <summary>
        ''' Get the header for specific partNumber and Message UID
        ''' </summary>
        ''' <param name="sMessageUid">Unique Identifier of message</param>
        ''' <param name="sPartNumber"> Message part number</param>
        ''' <param name="asMessageHeader">Return array </param>
        Sub GetHeader(sMessageUid As String, sPartNumber As String, asMessageHeader As ArrayList)
            asMessageHeader.Clear()
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandSuffix As String = ""
            If sPartNumber = "0" Then
                'sCommandSuffix = sMessageUid + " " + "BODY[TEXT]";//MCO
                sCommandSuffix = sMessageUid + " " + "BODY.PEEK[HEADER.FIELDS (SUBJECT)]"
            Else
                sCommandSuffix = sMessageUid + " " + "BODY[" + sPartNumber + ".MIME]"
            End If
            'string sCommandString = IMAP_UIDFETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL;
            Dim sCommandString As String = IMAP_FETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL
            'MCO
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    '-------------------------
                    ' PARSE RESPONSE
                    Dim sKey As String = ""
                    Dim sValue As String = ""
                    Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                    For Each sLine As String In asResultArray
                        If sLine.Length <= 0 OrElse sLine.StartsWith(IMAP_UNTAGGED_RESPONSE_PREFIX) OrElse sLine = ")" Then
                            Continue For
                        ElseIf sLine.StartsWith(sLastLine) Then

                            Exit For
                        End If
                        Dim nPos As Integer = sLine.IndexOf(" ")
                        If nPos <> -1 Then
                            Dim sTmpLine As String = sLine.Substring(0, nPos)
                            nPos = sTmpLine.IndexOf(":")
                        End If
                        If nPos <> -1 Then
                            If sKey.Length > 0 Then
                                asMessageHeader.Add(sKey)
                                asMessageHeader.Add(sValue)
                            End If
                            sKey = sLine.Substring(0, nPos).Trim()
                            sValue = sLine.Substring(nPos + 1).Trim()
                        Else
                            sValue += sLine.Trim()
                        End If
                    Next
                    If sKey.Length > 0 Then
                        asMessageHeader.Add(sKey)
                        asMessageHeader.Add(sValue)

                    End If
                Else
                    Throw New ImapException(ImapException.ImapErrorEnum.IMAP_ERR_FETCHMSG, sCommandSuffix)
                End If
            Catch e As ImapException
                LogOut()
                Throw e
            End Try


        End Sub
        ''' <summary>
        ''' Check if this message is multipart
        ''' To Identify multipart message, the content-type is either
        ''' multipart or rfc822
        ''' </summary>
        ''' <param name="asHeader"></param>
        ''' <returns></returns>
        Public Function IsMultipart(asHeader As ArrayList) As Boolean
            Dim i As Integer = 0
            While i < asHeader.Count
                If asHeader(i).ToString().ToLower() = IMAP_MESSAGE_CONTENT_TYPE.ToLower() Then
                    Dim sValue As String = asHeader(i + 1).ToString()
                    sValue = sValue.ToLower()
                    If sValue.StartsWith(IMAP_MESSAGE_MULTIPART.ToLower()) OrElse sValue.StartsWith(IMAP_MESSAGE_RFC822.ToLower()) Then
                        Return True
                    Else
                        Return False
                    End If
                End If
                i = i + 2
            End While
            Return False
        End Function

        ''' <summary>
        ''' Returns true if starts with NIL
        ''' </summary>
        ''' <param name="sBodyStruct">Body Structure</param>
        ''' <returns>true/false</returns>
        Function IsNilString(ByRef sBodyStruct As String) As Boolean
            Dim sSub As String = sBodyStruct.Substring(0, 3)
            If sSub = IMAP_MESSAGE_NIL Then
                sBodyStruct = sBodyStruct.Substring(3)
                Return True
            End If
            Return False
        End Function
        ''' <summary>
        ''' Get the content type
        ''' </summary>
        ''' <param name="sBodyStruct">Body Structure</param>
        ''' <param name="sType">Content Type</param>
        ''' <param name="sSubType">Sub Type</param>
        ''' <param name="sContentType">Content Type Value</param>
        ''' <returns>True/false</returns>
        Function GetContentType(ByRef sBodyStruct As String, ByRef sType As String, ByRef sSubType As String, ByRef sContentType As String) As Boolean
            sContentType = IMAP_PLAIN_TEXT

            ' get the type and the subtype strings from body struct
            ' If not found, set it to the default value plain/text.

            If sType.Length < 1 Then
                If Not ParseQuotedString(sBodyStruct, sType) Then
                    Log(LogTypeEnum.[ERROR], "Failed getting Content-Type.")
                    Return False
                End If
            End If
            sSubType = ""
            If Not ParseQuotedString(sBodyStruct, sSubType) Then
                Log(LogTypeEnum.[ERROR], "Failed getting Content-Sub-Type.")
                Return False
            End If

            If sType.Length > 0 AndAlso sSubType.Length > 0 Then
                sContentType = sType + "/" + sSubType
            End If

            ' Add extra parameters (if any) to the Content-type
            Dim asParam As New ArrayList()
            If Not ParseParameters(sBodyStruct, asParam) Then
                Log(LogTypeEnum.[ERROR], "Failed getting Content-Type Parameters.")
                Return False
            End If
            Dim i As Integer = 0
            While i < asParam.Count
                Dim sTemp As String = "; " + asParam(i).ToString() + "=""" + asParam(i + 1).ToString() + """"
                sContentType += sTemp
                i += 2
            End While
            Return True
        End Function
        ''' <summary>
        ''' Get Content Disposition
        ''' </summary>
        ''' <param name="sBodyStruct"> Body Structure</param>
        ''' <param name="sDisp">Content Disposition</param>
        ''' <returns>true if success</returns>
        Function GetContentDisposition(ByRef sBodyStruct As String, ByRef sDisp As String) As Boolean
            sDisp = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' find and remove opening paranthesis
            If Not FindAndRemove(sBodyStruct, "("c) Then
                Return False
            End If

            ' get the content disposition type (inline/attachment)
            If Not ParseQuotedString(sBodyStruct, sDisp) Then
                Log(LogTypeEnum.[ERROR], "Failed getting Content Disposition.")
                Return False
            End If
            ' get the associated parameters if any
            Dim asParam As New ArrayList()
            If Not ParseParameters(sBodyStruct, asParam) Then
                Log(LogTypeEnum.[ERROR], "Failed getting Content Disposition params.")
                Return False
            End If
            ' prepare the content-disposition
            Dim sTemp As String = ""
            Dim i As Integer = 0
            While i < asParam.Count
                sTemp = "; " + asParam(i).ToString() + "=""" + asParam(i + 1).ToString() + """"
                sDisp += sTemp
                i += 2
            End While
            ' remove the closing paranthesis
            Return FindAndRemove(sBodyStruct, ")"c)
        End Function
        ''' <summary>
        ''' Parse the quoted string in body structure
        ''' </summary>
        ''' <param name="sBodyStruct">Body Structure</param>
        ''' <param name="sString">"Quoted string</param>
        ''' <returns></returns>
        Function ParseQuotedString(ByRef sBodyStruct As String, ByRef sString As String) As Boolean
            sString = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If



            If sBodyStruct(0) = """"c Then
                ' extract the part within quotes
                Dim ch As Char
                Dim nEnd As Integer
                nEnd = 1

                'for (nEnd = 1; (ch = sBodyStruct[nEnd]) != '"'; nEnd++) 
                '				{
                '					if (ch == '\\') 
                '						nEnd++;
                '				}


                While (ch = sBodyStruct(nEnd)) <> """"""
                    If ch = "\"c Then
                        nEnd += 1
                    End If
                    nEnd += 1
                End While
                sString = sBodyStruct.Substring(1, nEnd - 1)
                sBodyStruct = sBodyStruct.Substring(nEnd + 1)
                Return True
            End If
            Dim sLog As String = "InValid Body Structure " + sBodyStruct + "."
            Log(LogTypeEnum.[ERROR], sLog)
            Return False
        End Function
        ''' <summary>
        ''' Parse the string (seperated by spaces or parenthesis)
        ''' </summary>
        ''' <param name="sBodyStruct">Body Struct</param>
        ''' <param name="sString">string</param>
        ''' <returns></returns>
        Function ParseString(ByRef sBodyStruct As String, ByRef sString As String) As Boolean
            sString = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()
            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' extract the literal as whole looking for a
            ' space or closing paranthesis character
            Dim ch As Char
            Dim nEnd As Integer, nLen As Integer
            nLen = sBodyStruct.Length

            nEnd = 0
            While nEnd < nLen
                ch = sBodyStruct(nEnd)

                If ch = " "c OrElse ch = ")"c Then
                    Exit While
                End If
                nEnd += 1
            End While

            If nEnd > 0 Then
                sString = sBodyStruct.Substring(0, nEnd)
                sBodyStruct = sBodyStruct.Substring(nEnd)
            End If
            Return True
        End Function
        ''' <summary>
        ''' Parse the language or list of languages in body structure
        ''' </summary>
        ''' <param name="sBodyStruct">Bosy structure</param>
        ''' <param name="sString">Languages</param>
        ''' <returns>true if success</returns>

        Function ParseLanguage(ByRef sBodyStruct As String, ByRef sString As String) As Boolean
            sString = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            If sBodyStruct(0) = "("c Then
                ' Language list
                ' remove the opening paranthesis
                If Not FindAndRemove(sBodyStruct, "("c) Then
                    Return False
                End If

                ' TO DO
                ' Logic for parsing language list is not yet
                ' written. To be added in the future.

                ' remove the closing paranthesis
                If Not FindAndRemove(sBodyStruct, ")"c) Then
                    Return False
                End If
            Else
                ' One or no language
                If Not ParseQuotedString(sBodyStruct, sString) Then
                    Log(LogTypeEnum.[ERROR], "Failed getting Content Language.")
                    Return False
                End If
            End If

            Return True
        End Function
        ''' <summary>
        ''' Parse the parameter in body structure
        ''' </summary>
        ''' <param name="sBodyStruct">Body structure</param>
        ''' <param name="asParams">parameter</param>
        ''' <returns>true if success</returns>
        Function ParseParameters(ByRef sBodyStruct As String, asParams As ArrayList) As Boolean
            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' remove the opening paranthesis
            If Not FindAndRemove(sBodyStruct, "("c) Then
                Return False
            End If

            Dim sName As String = ""
            Dim sValue As String = ""
            While sBodyStruct(0) <> ")"c

                ' Name
                If Not ParseQuotedString(sBodyStruct, sName) Then
                    Log(LogTypeEnum.[ERROR], "Invalid Body Parameter Name.")
                    Return False
                End If

                ' Value
                If Not ParseQuotedString(sBodyStruct, sValue) Then
                    Log(LogTypeEnum.[ERROR], "Invalid Body Parameter Value.")
                    Return False
                End If
                asParams.Add(sName)
                asParams.Add(sValue)
            End While

            ' remove the closing paranthesis
            Return FindAndRemove(sBodyStruct, ")"c)
        End Function
        ''' <summary>
        ''' Parse the extension in body structure
        ''' </summary>
        ''' <param name="sBodyStruct">body structure</param>
        ''' <param name="sString">extension</param>
        ''' <returns>true if success</returns>
        Function ParseExtension(ByRef sBodyStruct As String, ByRef sString As String) As Boolean
            sString = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' TO DO
            ' Dont know what to do with the data.

            Return True
        End Function
        ''' <summary>
        ''' Parse the address string
        ''' </summary>
        ''' <param name="sBodyStruct">body structure</param>
        ''' <param name="sString">address</param>
        ''' <returns>true if success</returns>
        Function ParseAddressList(ByRef sBodyStruct As String, ByRef sString As String) As Boolean
            sString = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' remove the opening paranthesis
            If Not FindAndRemove(sBodyStruct, "("c) Then
                Return False
            End If

            ' process each address
            Dim sAddress As String = ""
            While sBodyStruct(0) = "("c

                ' Get each address in the list
                If Not ParseAddress(sBodyStruct, sAddress) Then
                    Return True
                End If

                ' prepare the address list (as comma separated
                ' list of addresses).
                If sAddress.Length > 0 Then
                    If sString.Length > 0 Then
                        sString += ", "
                    End If
                    sString += sAddress
                End If

                sBodyStruct = sBodyStruct.Trim()
            End While

            ' remove the closing paranthesis
            Return FindAndRemove(sBodyStruct, ")"c)
        End Function
        ''' <summary>
        ''' Parse one address and format the string
        ''' </summary>
        ''' <param name="sBodyStruct">body structure</param>
        ''' <param name="sString">address</param>
        ''' <returns>true if success</returns>
        Function ParseAddress(ByRef sBodyStruct As String, ByRef sString As String) As Boolean
            sString = ""

            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' remove the opening paranthesis
            If Not FindAndRemove(sBodyStruct, "("c) Then
                Return False
            End If

            Dim sPersonal As String = ""
            Dim sEmailId As String = ""
            Dim sEmailDomain As String = ""
            Dim sTemp As String = ""

            ' Personal Name
            If Not ParseQuotedString(sBodyStruct, sPersonal) Then
                Log(LogTypeEnum.[ERROR], "Failed getting the Personal Name.")
                Return False
            End If
            ' At Domain List (Right now, don't know what to do with this)
            If Not ParseQuotedString(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Failed getting the Domain List.")
                Return False
            End If
            ' Email Id
            If Not ParseQuotedString(sBodyStruct, sEmailId) Then
                Log(LogTypeEnum.[ERROR], "Failed getting the Email Id.")
                Return False
            End If
            ' Email Domain
            If Not ParseQuotedString(sBodyStruct, sEmailDomain) Then
                Log(LogTypeEnum.[ERROR], "Failed getting the Email Domain.")
                Return False
            End If

            If sEmailId.Length > 0 Then
                If sPersonal.Length > 0 Then
                    If sEmailDomain.Length > 0 Then
                        sString = sPersonal + " <" + sEmailId + "@" + sEmailDomain + ">"
                    Else
                        sString = sPersonal + " <" + sEmailId + ">"
                    End If
                Else
                    If sEmailDomain.Length > 0 Then
                        sString = sEmailId + "@" + sEmailDomain
                    Else
                        sString = sEmailId
                    End If
                End If
            End If

            ' remove the closing paranthesis
            Return FindAndRemove(sBodyStruct, ")"c)
        End Function

        Function ParseEnvelope(ByRef sBodyStruct As String, asAttrs As ArrayList) As Boolean
            ' remove any spaces at the beginning and the end
            sBodyStruct = sBodyStruct.Trim()

            ' check if this is NIL
            If IsNilString(sBodyStruct) Then
                Return True
            End If

            ' look for '(' character
            If Not FindAndRemove(sBodyStruct, "("c) Then
                Return False
            End If

            Dim sTemp As String = ""
            ' Date
            If Not ParseQuotedString(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope Date.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_DATE_TAG)
                asAttrs.Add(sTemp)
            End If

            ' Subject
            If Not ParseQuotedString(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope Subject.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_SUBJECT_TAG)
                asAttrs.Add(sTemp)
            End If

            ' From
            If Not ParseAddressList(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.IMAP, "Invalid Message Envelope From.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_FROM_TAG)
                asAttrs.Add(sTemp)
            End If

            ' Sender
            If Not ParseAddressList(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope Sender.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_SENDER_TAG)
                asAttrs.Add(sTemp)
            End If

            ' ReplyTo
            If Not ParseAddressList(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope Reply-To.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_REPLY_TO_TAG)
                asAttrs.Add(sTemp)
            End If

            ' To
            If Not ParseAddressList(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.IMAP, "Invalid Message Envelope To.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_TO_TAG)
                asAttrs.Add(sTemp)
            End If

            ' Cc
            If Not ParseAddressList(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope CC.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_CC_TAG)
                asAttrs.Add(sTemp)
            End If

            ' Bcc
            If Not ParseAddressList(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope BCC.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_BCC_TAG)
                asAttrs.Add(sTemp)
            End If

            ' In-Reply-To
            If Not ParseQuotedString(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope In-Reply-To.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_HEADER_IN_REPLY_TO_TAG)
                asAttrs.Add(sTemp)
            End If

            ' Message Id
            If Not ParseQuotedString(sBodyStruct, sTemp) Then
                Log(LogTypeEnum.[ERROR], "Invalid Message Envelope Message Id.")
                Return False
            End If
            If sTemp.Length > 0 Then
                asAttrs.Add(IMAP_MESSAGE_ID)
                asAttrs.Add(sTemp)
            End If

            ' remove the closing paranthesis
            Return FindAndRemove(sBodyStruct, ")"c)
        End Function
        ''' <summary>
        ''' find the given character and remove
        ''' </summary>
        ''' <param name="sBodyStruct">body structure</param>
        ''' <param name="ch">first character to find and remove</param>
        ''' <returns>true if success</returns>
        Function FindAndRemove(ByRef sBodyStruct As String, ch As Char) As Boolean
            sBodyStruct = sBodyStruct.Trim()
            If sBodyStruct(0) <> ch Then
                Dim sLog As String = "Invalid Body Structure " + sBodyStruct + "."
                Log(LogTypeEnum.[ERROR], sLog)
                Return False
            End If

            ' remove character
            sBodyStruct = sBodyStruct.Substring(1)

            Return True
        End Function
        ''' <summary>
        ''' Get the Size of the fetch command response
        ''' response will look like "{<size>}"
        ''' </summary>
        ''' <param name="sResponse"></param>
        ''' <returns></returns>
        'long GetResponseSize(string sResponse)
        '{
        '    // check if there is a size element
        '    // if not, then the message part number is wrong

        '    if (sResponse.IndexOf(IMAP_MESSAGE_NIL) != -1) 
        '    {
        '        Log(LogTypeEnum.ERROR, "Size 0. No Message after this.");
        '        return 0L;
        '    }


        '    int nStart = sResponse.IndexOf('{');
        '    if (nStart == -1) 
        '    {
        '        string sLog = "Invalid size in Response " + sResponse + ".";
        '        Log(LogTypeEnum.ERROR,sLog);
        '        throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDHEADER);
        '    }
        '    int nEnd = sResponse.IndexOf('}');
        '    if (nEnd == -1 || nEnd < nStart) 
        '    {
        '        string sLog = "Invalid size in Response " + sResponse + ".";
        '        Log(LogTypeEnum.ERROR,sLog);
        '        throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDHEADER);
        '    }

        '    string sSize = sResponse.Substring( nStart+1, (nEnd - nStart -1));
        '    long nSize = Convert.ToInt64(sSize, 10 );

        '    if ( nSize <= 0L ) 
        '    {
        '        string sLog = "Invalid size in Response " + sResponse + ".";
        '        Log(LogTypeEnum.ERROR,sLog);
        '        throw new ImapException(ImapException.ImapErrorEnum.IMAP_ERR_INVALIDHEADER);
        '    }

        '    return nSize;
        '}

        Public Function isConnected() As [Boolean]
            Return m_bIsConnected
        End Function

        Public Function isLogged() As [Boolean]
            Return m_bIsLoggedIn
        End Function

        ''' <summary>
        ''' Nom de messages dans le dossier
        ''' </summary>
        Public ReadOnly Property nNbreMsgTotal() As Integer
            Get
                Return m_nTotalMessages
            End Get
        End Property

        ''' <summary>
        ''' Récupération du sujet d'un message
        ''' </summary>
        ''' <param name="sMessageUid"></param>
        ''' <returns></returns>
        Public Function GetSubject(sMessageUid As String) As [String]
            Dim sReturn As String = ""
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandSuffix As String = ""
            sCommandSuffix = sMessageUid + " " + "BODY[HEADER.FIELDS (SUBJECT)]"
            Dim sCommandString As String = IMAP_FETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL
            'MCO
            '-----------------------
            ' SEND SEARCH REQUEST
            Dim asResultArray As New ArrayList()
            eImapResponse = SendAndReceive(sCommandString, asResultArray)
            If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                '-------------------------
                ' PARSE RESPONSE
                Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                For Each sLine As String In asResultArray
                    If sLine.StartsWith("Subject:") Then
                        'Conversion de la chaine QuotedString to String
                        'Attachment attachment = Attachment.CreateAttachmentFromString("", sReturn);
                        'sReturn = attachment.Name;

                        sReturn = sLine.Replace("Subject:", "")

                    End If

                Next
            Else
                sReturn = ""
            End If

            Return sReturn.Trim()

        End Function
        ' GetSubject
        ''' <summary>
        ''' Get the Body structure of the message.
        ''' If message is single part then first part is 1
        ''' If message is multipart then first part is 0
        ''' </summary>
        ''' <param name="sMessageUID"> Message UID</param>
        Public Function GetBody(sMessageUID As String) As String
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            '            Dim _sCommandSuffix As String = sMessageUID + " " + "BODY"
            '            Dim _sCommandString As String = IMAP_FETCH_COMMAND + " " + _sCommandSuffix + IMAP_COMMAND_EOL
            Dim sCommandSuffix As String = sMessageUID + " " + "BODY[TEXT]"
            Dim sCommandString As String = IMAP_FETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL
            Dim sReturn As String = ""
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                '               Dim _asResultArray As New ArrayList()
                '              eImapResponse = SendAndReceive(_sCommandString, _asResultArray)
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                    Dim bCmd As Boolean = False
                    ' on va rechercher dans la réponse les balises [?xml] et [/xml] qui entoure la description de la commande
                    For Each sLine As String In asResultArray
                        If sLine.StartsWith("[?xml") Then
                            bCmd = True
                        End If
                        'La balise </xml> ne sera donc pas enregistré dans le résulat final !!!!
                        ' et c'est tant mieux car il n'est pas valide
                        If sLine.StartsWith("[/xml") Then
                            bCmd = False
                        End If
                        If bCmd Then
                            'Conversion de la chaine QuotedString to String
                            'Dim att As Attachment = Attachment.CreateAttachmentFromString("", sLine)
                            'sLine = att.Name
                            If sLine.EndsWith("=") Then
                                sLine = Left(sLine, sLine.Length - 1)
                            End If
                            sReturn += sLine
                        End If
                    Next
                    ' on remplace les [] par les <> traditionnels XML
                    sReturn = sReturn.Replace("=3D", "=")
                    sReturn = sReturn.Replace("[", "<")
                    sReturn = sReturn.Replace("]", ">")
                    sReturn = sReturn.Replace("#", "=")
                    sReturn = DecodequotedprintableString(sReturn, "utf-8")

                End If
            Catch e As ImapException
                LogOut()
                Throw e
            End Try
            Return sReturn
        End Function
        '
        ''' <summary>
        '''  Decode a quotedPrintable String
        ''' </summary>
        ''' <param name="pStrIn"></param>
        ''' <param name="encoding"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DecodequotedprintableString(pStrIn As String, encoding As String) As String
            Dim strOut As String = pStrIn.Replace("=" & vbCr & vbLf, "")
            strOut = strOut.Replace("=" & vbCrLf, "")
            strOut = strOut.Replace("=" & vbLf, "")
            strOut = strOut.Replace("=" & vbCr, "")
            Try
                ' Find the first =
                Dim position As Integer = strOut.IndexOf("=")
                While position <> -1
                    ' String before the =
                    Dim leftpart As String = strOut.Substring(0, position)
                    ' get the QuotedPrintable String in a ArrayList
                    Dim hex As New System.Collections.ArrayList()
                    ' The first Part
                    hex.Add(strOut.Substring(1 + position, 2))
                    ' Look for the next parts
                    While position + 3 < strOut.Length AndAlso strOut.Substring(position + 3, 1) = "="
                        'The Second Part
                        position = position + 3
                        hex.Add(strOut.Substring(1 + position, 2))
                    End While
                    ' In the hex Array, we have two items 
                    ' Convert using the GetEncoding Function
                    Dim equivalent As String
                    Dim nextPosition As Integer
                    Try
                        Dim bytes As Byte() = New Byte(hex.Count - 1) {}
                        For i As Integer = 0 To hex.Count - 1
                            bytes(i) = System.Convert.ToByte(New String(DirectCast(hex(i), String).ToCharArray()), 16)
                        Next
                        equivalent = System.Text.Encoding.GetEncoding(encoding).GetString(bytes)
                        nextPosition = position + 3
                    Catch ex As Exception
                        equivalent = "="
                        nextPosition = position + 1
                    End Try

                    ' Part of the orignal String after the last QP Symbol
                    Dim rightpart As String = strOut.Substring(nextPosition)
                    ' Re build the String
                    strOut = leftpart & equivalent & rightpart
                    ' find the new QP Position
                    position = leftpart.Length + equivalent.Length
                    If rightpart.Length = 0 Then
                        position = -1
                    Else
                        position = strOut.IndexOf("=", position + 1)
                    End If
                End While
            Catch ex As Exception

            End Try
            Return strOut
        End Function

        Public Function GetFromAdress(sMessageUID As String) As String
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandSuffix As String = sMessageUID + " " + "BODY[HEADER.FIELDS (FROM)]"
            Dim sCommandString As String = IMAP_FETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL
            Dim sReturn As String = ""
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                    Dim bCmd As Boolean = False
                    ' on va rechercher dans la réponse le mot from
                    For Each sLine As String In asResultArray
                        If sLine.StartsWith("From:") Then
                            sReturn = sLine.Replace("From:", "")
                        End If
                    Next

                End If
            Catch e As ImapException
                sReturn = ""
            End Try
            Return sReturn
        End Function

        Public Function GetToAdress(sMessageUID As String) As String
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandSuffix As String = sMessageUID + " " + "BODY[HEADER.FIELDS (TO)]"
            Dim sCommandString As String = IMAP_FETCH_COMMAND + " " + sCommandSuffix + IMAP_COMMAND_EOL
            Dim sReturn As String = ""
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                    ' on va rechercher dans la réponse le mot from
                    For Each sLine As String In asResultArray
                        If sLine.StartsWith("To:") Then
                            sReturn = sLine.Replace("To:", "")
                        End If
                    Next

                End If
            Catch e As ImapException
                sReturn = ""
            End Try
            Return sReturn
        End Function

        ''' <summary>
        ''' Récupération d'un message à partir de son uid
        ''' on ne récupére que le From, le To , le Sujet et le Body au format Text
        ''' </summary>
        ''' <param name="suid"></param>
        ''' <returns></returns>
        Public Function getMessage(suid As String) As System.Net.Mail.MailMessage
            Dim oMsg As New MailMessage()
            oMsg.Subject = GetSubject(suid)
            oMsg.Body = GetBody(suid)
            If Not String.IsNullOrEmpty(oMsg.Body) Then
                oMsg.From = New MailAddress(GetFromAdress(suid))
                oMsg.[To].Add(New MailAddress(GetToAdress(suid)))
            End If
            Return oMsg
        End Function
        ''' <summary>
        ''' Création d'un dossier à la racine de la mailBox
        ''' </summary>
        ''' <param name="Folder"></param>
        ''' <returns>OK,ALLREADYEXISTS</returns>
        ''' <remarks></remarks>
        Public Function createFolder(Folder As String) As String
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            '            Dim sCommandSuffix As String = " " + "LIST " + """" + """" + " " + """" + """"
            Dim sCommandSuffix As String = "CREATE " + """" + Folder + """"
            Dim sCommandString As String = sCommandSuffix + IMAP_COMMAND_EOL
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Return "OK"
                End If
                If eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE Then
                    Dim bAlreadyExists As Boolean = False
                    For Each sLine As String In asResultArray
                        If (sLine.IndexOf("ALREADYEXISTS") > 0) Then
                            bAlreadyExists = True
                        End If
                    Next
                    If bAlreadyExists Then
                        Return "ALREADYEXISTS"
                    Else
                        Return "NOK"
                    End If
                End If
            Catch e As ImapException
                Return "ERROR"
            End Try
        End Function
        Public Function deleteFolder(Folder As String) As String
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            '            Dim sCommandSuffix As String = " " + "LIST " + """" + """" + " " + """" + """"
            Dim sCommandSuffix As String = "DELETE " + """" + Folder + """"
            Dim sCommandString As String = sCommandSuffix + IMAP_COMMAND_EOL
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Return "OK"
                End If
                If eImapResponse = ImapResponseEnum.IMAP_FAILURE_RESPONSE Then
                    Dim bNoneExistent As Boolean = False
                    For Each sLine As String In asResultArray
                        If (sLine.IndexOf("NONEXISTENT") > 0) Then
                            bNoneExistent = True
                        End If
                    Next
                    If bNoneExistent Then
                        Return "NONEXISTENT"
                    Else
                        Return "NOK"
                    End If
                End If
            Catch e As ImapException
                Return "ERROR"
            End Try
        End Function
        Public Function ListFolder() As List(Of String)
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandSuffix As String = "LIST " + """" + """" + " " + """*" + """"
            Dim sCommandString As String = sCommandSuffix + IMAP_COMMAND_EOL
            Dim sReturn As New List(Of String)
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                    ' on va rechercher dans la réponse le mot from
                    For Each sLine As String In asResultArray
                        If sLine.StartsWith("To:") Then
                            sReturn.Add(sLine.Replace("To:", ""))
                        End If
                    Next

                End If
            Catch e As ImapException
                sReturn.Clear()
            End Try
            Return sReturn
        End Function
        Public Function GetFlags(suid As String) As List(Of String)
            Dim eImapResponse As ImapResponseEnum = ImapResponseEnum.IMAP_SUCCESS_RESPONSE
            Dim sCommandSuffix As String = "FETCH " + suid + " " + "FLAGS"
            Dim sCommandString As String = sCommandSuffix + IMAP_COMMAND_EOL
            Dim sReturn As New List(Of String)
            Try
                '-----------------------
                ' SEND SEARCH REQUEST
                Dim asResultArray As New ArrayList()
                eImapResponse = SendAndReceive(sCommandString, asResultArray)
                If eImapResponse = ImapResponseEnum.IMAP_SUCCESS_RESPONSE Then
                    Dim sLastLine As String = IMAP_SERVER_RESPONSE_OK
                    ' on va rechercher dans la réponse le mot FETCH
                    Dim npos As Integer
                    Dim str As String
                    For Each sLine As String In asResultArray
                        If sLine.IndexOf("FETCH") > 0 Then
                            'Puis le mot FLAGS
                            npos = sLine.IndexOf("FLAGS")
                            'On récupére la fin de la phrase et on supprime les paranthèses fermantes
                            str = sLine.Substring(npos + "FLAGS".Length + 2)
                            str = str.Replace("))", "")
                            'on récupère un chaine du style "\answered \Seen \Deleted"
                            'Split sur le \
                            sReturn.AddRange(str.Split("\"c))
                        End If
                    Next

                End If
            Catch e As ImapException
                sReturn.Clear()
            End Try
            Return sReturn
        End Function

        Public Function isFlagged(suid As String, pFlagName As String) As Boolean
            Dim lstFlags As List(Of String)
            Dim bReturn As Boolean = False
            lstFlags = GetFlags(suid)
            For Each Str As String In lstFlags
                If Trim(Str).Equals(Trim(pFlagName)) Then
                    bReturn = True
                    Exit For
                End If
            Next
            Return bReturn
        End Function
        Public Function isAnswered(suid As String) As Boolean
            Return isFlagged(suid, "Answered")
        End Function

    End Class
End Namespace