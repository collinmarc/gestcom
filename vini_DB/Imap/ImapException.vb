Imports System
Namespace ImapVB
    ''' <summary>
    ''' Imap Exception class which implements Imap releted exceptions
    ''' </summary>
    Public Class ImapException
        Inherits ApplicationException
        ''' <summary>
        ''' Exception message string
        ''' </summary>
        Private m_message As String
        ''' <summary>
        ''' enum for Imap exception errors
        ''' </summary>
        Public Enum ImapErrorEnum
            ''' <summary>
            ''' failure parsing the url
            ''' </summary>
            IMAP_ERR_URI
            ''' <summary>
            ''' invalid message uid in the url
            ''' </summary>
            IMAP_ERR_MESSAGEUID
            ''' <summary>
            ''' invalid username/password in the url
            ''' </summary>
            IMAP_ERR_AUTHFAILED
            ''' <summary>
            ''' failure connecting to imap server
            ''' </summary>
            IMAP_ERR_CONNECT
            ''' <summary>
            ''' not connected to any IMAP
            ''' </summary>
            IMAP_ERR_NOTCONNECTED
            ''' <summary>
            ''' failure logging into imap server
            ''' </summary>
            IMAP_ERR_LOGIN
            ''' <summary>
            ''' failure to logout from imap server
            ''' </summary>
            IMAP_ERR_LOGOUT
            ''' <summary>
            ''' not enough data to restore
            ''' </summary>
            IMAP_ERR_INSUFFICIENT_DATA
            ''' <summary>
            ''' timeout while waiting for response
            ''' </summary>
            IMAP_ERR_TIMEOUT
            ''' <summary>
            ''' socket error while receiving
            ''' </summary>
            IMAP_ERR_SOCKET
            ''' <summary>
            ''' failure getting the quota information
            ''' </summary>
            IMAP_ERR_QUOTA
            ''' <summary>
            ''' failure selecting a IMAP folder
            ''' </summary>
            IMAP_ERR_SELECT
            ''' <summary>
            ''' failure examining an IMAP folder
            ''' </summary>
            IMAP_ERR_EXAMINE
            ''' <summary>
            '''  No folder is currently selected
            ''' </summary>
            IMAP_ERR_NOTSELECTED
            ''' <summary>
            ''' failure to search
            ''' </summary>
            IMAP_ERR_SEARCH
            ''' <summary>
            ''' failed to do exact match after search
            ''' </summary>
            IMAP_ERR_SEARCH_EXACT
            ''' <summary>
            ''' unsupported search key
            ''' </summary>
            IMAP_ERR_INVALIDSEARCHKEY
            ''' <summary>
            ''' failure to get message MIME
            ''' </summary>
            IMAP_ERR_GETMIME
            ''' <summary>
            ''' Message Header is in invalid format
            ''' </summary>
            IMAP_ERR_INVALIDHEADER
            ''' <summary>
            ''' Failed to fetch the bodystructure
            ''' </summary>
            IMAP_ERR_FETCHBODYSTRUCT
            ''' <summary>
            ''' failure to fetch a IMAP message
            ''' </summary>
            IMAP_ERR_FETCHMSG
            ''' <summary>
            ''' failure to fetch a IMAP message size
            ''' </summary>
            IMAP_ERR_FETCHSIZE
            ''' <summary>
            ''' failure to allocate memory
            ''' </summary>
            IMAP_ERR_MEMALLOC
            ''' <summary>
            ''' failure to encode the audio content
            ''' </summary>
            IMAP_ERR_ENCODINGERROR
            ''' <summary>
            ''' failure to read/write the audio content
            ''' </summary>
            IMAP_ERR_FILEIO
            ''' <summary>
            ''' failure to store the message in IMAP
            ''' </summary>
            IMAP_ERR_STOREMSG
            ''' <summary>
            ''' failure to issue expunge command
            ''' </summary>
            IMAP_ERR_EXPUNGE
            ''' <summary>
            ''' invalid parameter to API
            ''' </summary>
            IMAP_ERR_INVALIDPARAM
            ''' <summary>
            ''' Capability command error
            ''' </summary>
            IMAP_ERR_CAPABILITY
            ''' <summary>
            ''' Serious Problem
            ''' </summary>
            IMAP_ERR_SERIOUS

        End Enum
        ''' <summary>
        ''' Property Message (string)
        ''' </summary>
        Public Property Message As String
            Get
                Return m_message
            End Get
            Set(value As String)
                m_message = value
            End Set
        End Property

        ''' <summary>
        ''' Error Type: ImapErrorEnum
        ''' </summary>
        Private errorType As ImapErrorEnum
        ''' <summary>
        ''' Property : Type (ImapErrorEnum)
        ''' </summary>
        Public Property Type() As ImapErrorEnum
            Get
                Return errorType
            End Get
            Set(value As ImapErrorEnum)
                errorType = value
            End Set
        End Property

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="message">string</param>
        Public Sub New(message As [String])
            MyBase.New(message)
            Me.Message = message
        End Sub
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="message">string</param>
        ''' <param name="inner">Exception</param>
        Public Sub New(message As [String], inner As Exception)
            MyBase.New(message, inner)
            Me.Message = message
        End Sub
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="Type">ImapErrorEnum</param>
        Public Sub New(Type As ImapErrorEnum)
            errorType = Type
            Message = GetDescription(Type)
        End Sub
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="Type">ImapErrorEnum</param>
        ''' <param name="error">string</param>
        Public Sub New(Type As ImapErrorEnum, [error] As String)
            errorType = Type
            Message = GetDescription(Type)
            Message = Message + " " + [error]
        End Sub

        ''' <summary>
        ''' Get Description for specified Type
        ''' </summary>
        ''' <param name="Type">ImapErrorEnum type</param>
        ''' <returns>string</returns>
        Private Function GetDescription(Type As ImapErrorEnum) As String
            Select Case Type
                Case ImapErrorEnum.IMAP_ERR_URI
                    Return "Failure parsing the IMAP URL."
                Case ImapErrorEnum.IMAP_ERR_MESSAGEUID
                    Return "Invalid UserName/Password in the IMAP URL."
                Case ImapErrorEnum.IMAP_ERR_CONNECT
                    Return "Failure connecting to the IMAP server."
                Case ImapErrorEnum.IMAP_ERR_NOTCONNECTED
                    Return "Not connected to any IMAP server. Do Login first."
                Case ImapErrorEnum.IMAP_ERR_LOGIN
                    Return "Failed authenticating the user/password in the IMAP server."
                Case ImapErrorEnum.IMAP_ERR_INSUFFICIENT_DATA
                    Return "Not enough data to Restore the connection."
                Case ImapErrorEnum.IMAP_ERR_LOGOUT
                    Return "Failed to Logout from IMAP server."
                Case ImapErrorEnum.IMAP_ERR_TIMEOUT
                    Return "Timeout while waiting for response from the IMAP server."
                Case ImapErrorEnum.IMAP_ERR_SOCKET
                    Return "Socket Error. Failed receiving message."
                Case ImapErrorEnum.IMAP_ERR_QUOTA
                    Return "Failure getting the quota information for the folder/mailbox."
                Case ImapErrorEnum.IMAP_ERR_SELECT
                    Return "Failure selecting IMAP folder/mailbox."
                Case ImapErrorEnum.IMAP_ERR_NOTSELECTED
                    Return "No Folder is currently selected."
                Case ImapErrorEnum.IMAP_ERR_EXAMINE
                    Return "Failure examining IMAP folder/mailbox."
                Case ImapErrorEnum.IMAP_ERR_SEARCH
                    Return "Failure searching IMAP with the given criteria."
                Case ImapErrorEnum.IMAP_ERR_SEARCH_EXACT
                    Return "Failure getting the exact search value."
                Case ImapErrorEnum.IMAP_ERR_INVALIDSEARCHKEY
                    Return "Unsupported search key passed to SearchMessage API."
                Case ImapErrorEnum.IMAP_ERR_GETMIME
                    Return "Failure fetching mime for the message."
                Case ImapErrorEnum.IMAP_ERR_FETCHMSG
                    Return "Failure fetching message from IMAP folder/mailbox."
                Case ImapErrorEnum.IMAP_ERR_FETCHSIZE
                    Return "Failure fetching message size from IMAP folder/mailbox."
                Case ImapErrorEnum.IMAP_ERR_MEMALLOC
                    Return "Failure allocating memory."
                Case ImapErrorEnum.IMAP_ERR_ENCODINGERROR
                    Return "Failure encoding audio message."
                Case ImapErrorEnum.IMAP_ERR_FILEIO
                    Return "Failure reading/writing the audio content to file."
                Case ImapErrorEnum.IMAP_ERR_STOREMSG
                    Return "Failure storing message in IMAP."
                Case ImapErrorEnum.IMAP_ERR_EXPUNGE
                    Return "Failure permanently deleting marked messages (EXPUNGE)."
                Case ImapErrorEnum.IMAP_ERR_INVALIDPARAM
                    Return "Invalid parameter passed to API. See Log for details"
                Case ImapErrorEnum.IMAP_ERR_CAPABILITY
                    Return "Failure capability command."
                Case ImapErrorEnum.IMAP_ERR_SERIOUS
                    Return "Serious Problem with IMAP. Contact System Admin."
                Case ImapErrorEnum.IMAP_ERR_FETCHBODYSTRUCT
                    Return "Failure bodystructure command"
                Case Else
                    Throw New Exception("UnKnow Exception")

            End Select

        End Function
    End Class
End Namespace