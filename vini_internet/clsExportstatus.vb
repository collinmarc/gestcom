Public Class clsExportstatus

    Private m_date As Date
    Private m_Message As String
    Public Property statusDate() As Date
        Get
            Return m_date
        End Get
        Set(ByVal value As Date)
            m_date = value
        End Set
    End Property


    Public Property statusMessage() As String
        Get
            Return m_Message
        End Get
        Set(ByVal value As String)
            m_Message = value

        End Set
    End Property

End Class
