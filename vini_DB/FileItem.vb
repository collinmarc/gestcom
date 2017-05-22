Option Strict
Option Explicit

Imports System
Imports System.IO

'/// <summary>Represents a file from the local or a remote system.</summary>
Public Class FileItem
    Implements IComparable
    '/// <summary>Constructs a new FileItem object.</summary>
    Public Sub New()
        'Do Nothing
    End Sub
    '/// <summary>Constructs a new FileItem object.</summary>
    '/// <param name="ObjectName">Specifies the full path of the object in question.</param>
    '/// <param name="IsDirectory">Specifies whether the specified object is a directory or a file.</param>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specifed Filename is Nothing (C#, VC++: null)</exceptions>
    '/// <exceptions cref="ArgumentException">Thrown when there was an error querying the information of the specified object.</exceptions>
    Public Sub New(ByVal ObjectName As String, ByVal IsDirectory As Boolean)
        If ObjectName Is Nothing Then Throw New ArgumentNullException()
        m_IsDirectory = IsDirectory
        If m_IsDirectory Then
            Try
                Dim di As New DirectoryInfo(ObjectName)
                m_FileTitle = di.Name
                m_FileSize = 0
                m_FilePath = di.FullName
                m_FilePath = m_FilePath.Substring(0, m_FilePath.Length - m_FileTitle.Length)
                m_FileDate = di.LastWriteTime
            Catch
                Throw New ArgumentException()
            End Try
        Else
            Try
                Dim fi As New FileInfo(ObjectName)
                m_FileTitle = fi.Name
                m_FileSize = fi.Length
                m_FilePath = fi.DirectoryName
                m_FileDate = fi.LastWriteTime
            Catch
                Throw New ArgumentException()
            End Try
        End If
    End Sub
    '/// <summary>Specifies the path of the object.</summary>
    '/// <value>A String that specifies the path of the object.</value>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null).</exceptions>
    Public Property FilePath() As String
        Get
            Return m_FilePath
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then Throw New ArgumentNullException()
            m_FilePath = value
        End Set
    End Property
    '/// <summary>Specifies whether the object is directory or not.</summary>
    '/// <value>A Boolean that specifies whether the object is a directory or not.</value>
    Public Property IsDirectory() As Boolean
        Get
            Return m_IsDirectory
        End Get
        Set(ByVal value As Boolean)
            m_IsDirectory = value
        End Set
    End Property
    '/// <summary>Specifies the date of the object.</summary>
    '/// <value>A DateTime object that specifies the date of the object.</value>
    Public Property FileDate() As DateTime
        Get
            Return m_FileDate
        End Get
        Set(ByVal value As DateTime)
            m_FileDate = value
        End Set
    End Property
    '/// <summary>Specifies the permissions of the file.</summary>
    '/// <value>A String that specifies the permissions for the file.</value>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null).</exceptions>
    Public Property FilePermissions() As String
        Get
            Return m_FilePermissions
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then Throw New ArgumentNullException()
            m_FilePermissions = value
        End Set
    End Property
    '/// <summary>Specifies the group of users the file belongs to.</summary>
    '/// <value>A String that specifies the group of users the file belongs to.</value>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null).</exceptions>
    Public Property FileGroup() As String
        Get
            Return m_FileGroup
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then Throw New ArgumentNullException()
            m_FileGroup = value
        End Set
    End Property
    '/// <summary>Specifies the owner of the file.</summary>
    '/// <value>A String that specifies the owner of the file.</value>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null).</exceptions>
    Public Property FileOwner() As String
        Get
            Return m_FileOwner
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then Throw New ArgumentNullException()
            m_FileOwner = value
        End Set
    End Property
    '/// <summary>Specifies the size of the file.</summary>
    '/// <value>A Long that specifies the size of the file.</value>
    Public Property FileSize() As Long
        Get
            Return m_FileSize
        End Get
        Set(ByVal value As Long)
            m_FileSize = value
        End Set
    End Property
    '/// <summary>Specifies the title of the file.</summary>
    '/// <value>A String that specifies the title of the file.</value>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null).</exceptions>
    Public Property FileTitle() As String
        Get
            Return m_FileTitle
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then Throw New ArgumentNullException()
            m_FileTitle = value
        End Set
    End Property
    '/// <summary>Compares this object to another FileItem object.</summary>
    '/// <returns>Returns 1 if the passed FileItem object should be placed above this FileItem, -1 if the passed FileItem should be placed below this FileItem and 0 if it is the same.</returns>
    Public Overridable Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        If obj Is Nothing Then Return -1
        Dim ct As FileItem = CType(obj, FileItem)
        If m_IsDirectory AndAlso Not (ct.IsDirectory) Then
            Return -1
        ElseIf Not (m_IsDirectory) AndAlso ct.IsDirectory Then
            Return 1
        ElseIf m_FileTitle.ToLower > ct.FileTitle.ToLower Then
            Return 1
        ElseIf m_FileTitle.ToLower < ct.FileTitle.ToLower Then
            Return -1
        Else
            Return 0
        End If
    End Function
    ' Private Variables
    Private m_FileTitle As String = ""
    Private m_FileSize As Long = 0
    Private m_FileOwner As String = ""
    Private m_FileGroup As String = ""
    Private m_FilePermissions As String = ""
    Private m_FileDate As DateTime
    Private m_IsDirectory As Boolean
    Private m_FilePath As String = ""
End Class
