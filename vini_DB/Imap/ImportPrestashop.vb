Imports System.Collections.Generic
Imports vini_DB.ImapVB
''' <summary>
''' Classe d'import des commandes depuis le Message Mail de prestashop
''' </summary>
''' <remarks></remarks>
Public Class ImportPrestashop
    Inherits racine
    Public Const IMPORTLOGFILE As String = "./Import.log"
    Private m_Hostname As String
    Private m_Port As Integer
    Private m_username As String
    Private m_password As String
    Private m_bSSL As Boolean
    Private m_nSec As Integer

    Private m_MSGFolderName As String

    Public Property HostName As String
        Get
            Return m_Hostname
        End Get
        Set(value As String)
            m_Hostname = value
        End Set
    End Property
    Public Property UserName As String
        Get
            Return m_username
        End Get
        Set(value As String)
            m_username = value
        End Set
    End Property
    Public Property Password As String
        Get
            Return m_password
        End Get
        Set(value As String)
            m_password = value
        End Set
    End Property
    Public Property MSGFolderName As String
        Get
            Return m_MSGFolderName
        End Get
        Set(value As String)
            m_MSGFolderName = value
        End Set
    End Property
    Public Property bSSL As Boolean
        Get
            Return m_bSSL
        End Get
        Set(value As Boolean)
            m_bSSL = value
        End Set
    End Property
    Public Property Port As Integer
        Get
            Return m_Port
        End Get
        Set(value As Integer)
            m_Port = value
        End Set
    End Property
    ''' <summary>
    ''' Nbre de secondes d'attentes entre 2 boucles (mode service uniquement)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Public Property nSec As Integer
    '    Get
    '        Return m_nSec
    '    End Get
    '    Set(value As Integer)
    '        m_nSec = value
    '    End Set
    'End Property

    Public Sub New(pHost As String, pUser As String, pPassword As String, Optional pPort As Integer = 993, Optional pSSL As Boolean = True)
        HostName = pHost
        UserName = pUser
        Password = pPassword
        Port = pPort
        bSSL = pSSL
        '        nSec = My.Settings.ImapNSec
    End Sub
    ''' <summary>
    ''' Initialisation par défaut à partir des informations  du .config
    ''' </summary>
    ''' <remarks></remarks>
    'Public Sub New()
    '    HostName = My.Settings.ImapHost
    '    UserName = My.Settings.ImapUser
    '    Password = My.Settings.ImapPassword
    '    Port = My.Settings.ImapPort
    '    bSSL = My.Settings.ImapSSL
    '    MSGFolderName = My.Settings.IMAPMSGfolder
    '    nSec = My.Settings.ImapNSec
    'End Sub

    Public Overrides Function toString() As String
        Return Me.GetType().ToString() & "(Host:[" & HostName & "],port:[" & Port & "],user:[" & UserName & "],Pwd:[" & Password & "]"
    End Function

    Public Overrides ReadOnly Property shortResume As String
        Get
            Return "ImportPrestashop " & toString()
        End Get
    End Property
    Public Function TestLogin() As Boolean
        Dim bReturn As Boolean
        Dim oImap As New Imap()
        bReturn = oImap.Login(HostName, Convert.ToUInt16(Port), UserName, Password, bSSL)
        If bReturn Then
            oImap.LogOut()
        End If
        Return bReturn
    End Function
    ''' <summary>
    ''' Mécanisme d'import de commande depuis la boite mail de réception Prestashop
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Import(Optional pbSaveCmd As Boolean = True, Optional pCheck As Boolean = True) As List(Of CommandeClient)
        Dim bReturn As Boolean
        Dim oReturn As New List(Of CommandeClient)
        Try
            Dim oImap As New Imap()
            oImap.Login(HostName, Convert.ToUInt16(Port), UserName, Password, bSSL)
            Dim oMsg As System.Net.Mail.MailMessage
            Dim oCmdCLT As CommandeClient
            'Création du Folder MSGTRAITE
            oImap.createFolder(MSGFolderName)
            oImap.SelectFolder("INBOX")
            For nmsg As Integer = 1 To oImap.nNbreMsgTotal
                oMsg = oImap.getMessage(nmsg)
                If Not String.IsNullOrEmpty(oMsg.Body) Then
                    'Création de la Commande ViniCom
                    oCmdCLT = createCMDCLT(oMsg, pCheck)
                    If oCmdCLT IsNot Nothing Then
                        oReturn.Add(oCmdCLT)
                        oImap.MoveMessage(nmsg, MSGFolderName)
                    End If
                End If
            Next
            oImap.SelectFolder("INBOX")
            For nmsg As Integer = oImap.nNbreMsgTotal To 1 Step -1
                If oImap.isAnswered(nmsg) Then
                    oImap.DeleteMessage2(nmsg)

                End If
            Next

            oImap.LogOut()

            If pbSaveCmd Then
                'Sauvegarde des commandes
                For Each oCmd As CommandeClient In oReturn
                    oCmd.save()
                Next
            End If
            bReturn = True
        Catch ex As Exception
            setError("ImportPrestashop", ex.Message)
            bReturn = False
        End Try
        Return oReturn
    End Function
    ''' <summary>
    ''' Création d'une commande Client à partir d'un mail
    ''' </summary>
    ''' <param name="pMsg">Message conteant la commande</param>
    ''' <param name="pCheck">Vérification préalable oui ou non (Oui par defaut)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function createCMDCLT(pMsg As System.Net.Mail.MailMessage, Optional pbCheck As Boolean = True) As CommandeClient
        Dim oReturn As CommandeClient
        Try
            Dim oCmd As cmdprestashop
            oReturn = Nothing
            oCmd = cmdprestashop.readXML(pMsg.Body)
            If oCmd IsNot Nothing Then
                Dim bCreate As Boolean
                If pbCheck Then
                    bCreate = oCmd.check()
                Else
                    bCreate = True
                End If
                If bCreate Then
                    oReturn = oCmd.createCommandeClient()
                    setError("  Commande N° " & oCmd.id & "(" & oCmd.name & ") importée")
                Else
                    setError("==Commande N° " & oCmd.id & "(" & oCmd.name & ") refusée : Motif " & oCmd.motif)
                End If
            End If
        Catch ex As Exception
            setError("ImportPrestashop.createCMDCLT", ex.Message)
            oReturn = Nothing
        End Try
        Return oReturn
    End Function

End Class
