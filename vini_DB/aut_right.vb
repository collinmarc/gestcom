Public Class aut_right
    Inherits Persist

    Private m_tag As String
    Private m_role As vncEnums.userRole
    Private m_text As String
    Private m_droit As Boolean
#Region "Accesseurs"
    Public Sub New(ByVal ptag As String, ByVal pRole As userRole, ByVal pRight As Boolean)
        MyClass.new()
        m_tag = ptag
        m_droit = pRight
        m_role = pRole
        m_text = ""
    End Sub
    Friend Sub New()
        m_role = vncEnums.userRole.INVITE
    End Sub

    Public Property tag() As String
        Get
            Return m_tag
        End Get
        Set(ByVal Value As String)
            m_tag = Value
        End Set
    End Property
    Public Property droit() As Boolean
        Get
            Return m_droit
        End Get
        Set(ByVal Value As Boolean)
            m_droit = Value
        End Set
    End Property

    Public Property role() As userRole
        Get
            Return m_role
        End Get
        Set(ByVal Value As userRole)
            m_role = Value
        End Set
    End Property
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_tag & "->" & m_droit
        End Get
    End Property

    Public Property text() As String
        Get
            Return m_text
        End Get
        Set(ByVal Value As String)
            m_text = Value
        End Set
    End Property

#End Region

    Public Overrides Function checkForDelete() As Boolean
        Return True
    End Function

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean

    End Function

    Friend Overrides Function delete() As Boolean

    End Function

    Friend Overrides Function insert() As Boolean
        Debug.Assert(id = 0, "idRight = 0")
        Dim bReturn As Boolean
        Try

            bReturn = insertRIGHT()
        Catch ex As Exception
            bReturn = False
            setError("au_right.insert", ex.ToString())
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function


    Friend Overrides Function update() As Boolean

    End Function
End Class
