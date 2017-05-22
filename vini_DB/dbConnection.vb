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
Imports System.data.OleDb
Imports System.Globalization
Imports System.Security.Permissions
Imports System.Threading
'<Assembly: SecurityPermission(SecurityAction.RequestMinimum, ControlThread:=True)> 
Public Class dbConnection
    Inherits racine

    'Private m_strprovider As String
    'Private m_strDataSource As String
    'Private m_ConnectionString As String
    Private m_dbconn As OleDbConnection
    Friend m_nConnection As Integer 'Nbre de Connection ouvertes
    Protected m_objTransaction As OleDbTransaction
    Public Shared Event Connected()
    Public Shared Event Disconnected()

    Friend Function connect(ByVal pstrConnectionString As String) As Boolean
        Dim bReturn As Boolean
        '   System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-GB") 

        'Dim oCulture As CultureInfo
        'oCulture = New System.Globalization.CultureInfo("fr-FR")
        'oCulture.NumberFormat.NumberDecimalSeparator = "."
        'System.Threading.Thread.CurrentThread.CurrentCulture = oCulture

        bReturn = True
        If Not isConnected() Then
            If (m_dbconn Is Nothing) Then
                m_dbconn = New OleDbConnection
            End If
            m_dbconn.ConnectionString = pstrConnectionString
            Try
                m_dbconn.Open()
                cleanErreur()
                m_nConnection = 1
                bReturn = True
            Catch ex As Exception
                setError("Shared_Connect", ex.ToString)
                bReturn = False
            End Try
        Else
            m_nConnection = m_nConnection + 1
        End If
        Return bReturn
    End Function ' Shared_Connect
    'Nom disconnect - Méthode de classe
    'Fonction Deconnection de la base de données
    Public Function disconnect() As Boolean
        Dim bReturn As Boolean
        bReturn = True
        If isConnected() Then
            Try
                'On ne déconnecte la base qui si le nombre de connection simultanée est à Zéro
                'CAD qui dans un enchainement de méthode on peut empiler les Connect et les disconnects
                'Seul le premier connect et le dernier disconnect sont important
                m_nConnection = m_nConnection - 1
                If m_nConnection < 1 Then
                    m_dbconn.Close()
                    m_nConnection = 0
                    bReturn = True
                End If
            Catch ex As Exception
                setError("SharedDisconnect", ex.ToString)
                bReturn = False
            End Try
        End If
        Debug.Assert(bReturn, getErreur())
        RaiseEvent Disconnected()
        Return bReturn
    End Function
    Public Function isConnected() As Boolean
        Return (m_nConnection > 0)
    End Function
    Public ReadOnly Property nConnection() As Integer
        Get
            Return m_nConnection
        End Get
    End Property
    Public ReadOnly Property Connection() As OleDbConnection
        Get
            Return m_dbconn
        End Get
    End Property
    'Friend Sub SaveConnection()
    '    m_strprovider2 = m_strprovider
    '    m_strDataSource2 = m_strDataSource
    '    m_ConnectionString2 = m_ConnectionString
    '    m_dbconn2 = m_dbconn
    '    m_nConnection2 = m_nConnection

    '    m_strprovider = ""
    '    m_strDataSource = ""
    '    m_ConnectionString = ""
    '    m_dbconn = Nothing
    '    m_nConnection = 0

    'End Sub
    'Friend Sub RestoreConnection()
    '    m_strprovider = m_strprovider2
    '    m_strDataSource = m_strDataSource2
    '    m_ConnectionString = m_ConnectionString2
    '    m_dbconn = m_dbconn2
    '    m_nConnection = m_nConnection2

    '    m_strprovider2 = ""
    '    m_strDataSource2 = ""
    '    m_ConnectionString2 = ""
    '    m_dbconn2 = Nothing
    '    m_nConnection2 = 0

    'End Sub
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "DBConnection =(" & ")"
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

    Public Sub New()
    End Sub

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
    Public Sub BeginTransaction()
        m_objTransaction = m_dbconn.BeginTransaction()
    End Sub
    Public ReadOnly Property transaction() As OleDbTransaction
        Get
            Return m_objTransaction
        End Get
    End Property


    Public Function Backup(ByVal strFichier As String) As Boolean

        Dim bReturn As Boolean
        Try
            Dim strDB As String
            Dim strSQL As String
            Dim objCommand As OleDbCommand
            strDB = m_dbconn.Database
            If m_dbconn.State = ConnectionState.Closed Then
                m_dbconn.Open()
            End If

            strSQL = "BACKUP DATABASE " + strDB + " TO DISK = '" + strFichier + "' WITH INIT"
            objCommand = New OleDbCommand(strSQL, m_dbconn)
            objCommand.CommandTimeout = 0
            objCommand.ExecuteNonQuery()
            bReturn = True
        Catch ex As Exception
            Persist.setError("dbconnection.backup", ex.Message)
            bReturn = False

        End Try
        Return bReturn
    End Function

    Public Function Restore(ByVal strFichier As String) As Boolean

        Dim bReturn As Boolean
        Try
            Dim strDB As String
            Dim strSQL As String
            Dim objCommand As OleDbCommand
            strDB = m_dbconn.Database
            If m_dbconn.State = ConnectionState.Closed Then
                m_dbconn.Open()
            End If

            strSQL = "use master; ALTER DATABASE " + strDB + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"
            objCommand = New OleDbCommand(strSQL, m_dbconn)

            objCommand.CommandTimeout = 0
            objCommand.ExecuteNonQuery()

            strSQL = " RESTORE DATABASE " + strDB + " FROM DISK = '" + strFichier + "' WITH REPLACE"
            objCommand = New OleDbCommand(strSQL, m_dbconn)
            objCommand.CommandTimeout = 0
            objCommand.ExecuteNonQuery()

            strSQL = "ALTER DATABASE " + strDB + " SET MULTI_USER; use " + strDB
            objCommand = New OleDbCommand(strSQL, m_dbconn)
            objCommand.CommandTimeout = 0
            objCommand.ExecuteNonQuery()
            bReturn = True
        Catch ex As Exception
            Persist.setError("dbconnection.restore", ex.Message)
            bReturn = False

        End Try
        Return bReturn
    End Function


End Class
