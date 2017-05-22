Public Class aut_user
    Inherits Persist


    Private m_code As String
    Private m_password As String
    Private m_role As userRole
    Private m_colRights As ColEvent
#Region "Accesseurs"
    Public Sub New(ByVal pCode As String, ByVal pPassword As String, ByVal pRole As userRole)
        MyClass.New()
        m_code = pCode
        m_role = pRole
        m_password = pPassword
        m_role = pRole
    End Sub
    Friend Sub New()
        m_colRights = New ColEvent
    End Sub

    Public Property code() As String
        Get
            Return m_code
        End Get
        Set(ByVal Value As String)
            If Value <> m_code Then
                RaiseUpdated()
                m_code = Value
            End If
        End Set
    End Property

    Public Property password() As String
        Get
            Return m_password
        End Get
        Set(ByVal Value As String)
            If Value <> m_password Then
                m_password = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public ReadOnly Property role() As userRole
        Get
            Return m_role
        End Get
        'Set(ByVal Value As userRole)
        '    If Value <> m_role Then
        '        m_role = Value
        '        RaiseUpdated()
        '    End If
        'End Set
    End Property
    Friend Sub setRole(ByVal prole As userRole)
        If prole <> m_role Then
            m_role = prole
            RaiseUpdated()
        End If
    End Sub
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_code
        End Get
    End Property

    Public ReadOnly Property colRigths() As ColEvent
        Get
            Return m_colRights
        End Get
    End Property

#End Region
#Region "Interface Persist"

    Public Overrides Function checkForDelete() As Boolean
        Return True
    End Function

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If
        Debug.Assert(id <> 0, "idUser <> 0")
        bReturn = LoadUSER()
        If m_id <> 0 Then
            bReturn = loadcolUSERSRIGHTS()
        End If
        Return bReturn
    End Function

    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("au_user.delete", ex.ToString())
        End Try

        Debug.Assert(bReturn, getErreur())
        Return False

    End Function

    Friend Overrides Function insert() As Boolean
        Debug.Assert(id = 0, "idUser = 0")
        Dim bReturn As Boolean
        Try

            bReturn = insertUSER()
            If bReturn Then
                bReturn = deleteColUSERSRIGHTS()
            End If
            If bReturn Then
                bReturn = insertcolUSERSRights()
            End If
 
        Catch ex As Exception
            bReturn = False
            setError("au_user.insert", ex.ToString())
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function


    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = updateUSER()
            If bReturn Then
                bReturn = deleteColUSERSRIGHTS()
            End If
            If bReturn Then
                bReturn = insertcolUSERSRights()
            End If
            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("au_user.update", ex.ToString())
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn

    End Function
#End Region
#Region "Méthodes de classes"
    Shared Function getuser(ByVal prole As userRole) As aut_user
        Dim col As Collection
        Dim objUser As aut_user = Nothing
        col = listeUSERS()
        If Not col Is Nothing Then
            For Each objUser In col
                If objUser.role = prole Then
                    Exit For
                End If
            Next
        End If
        Return objUser
    End Function
#End Region
#Region "Méthodes"
    Public Function ajouteDroit(ByVal ptag As String, ByVal pDroit As Boolean, ByVal pText As String) As aut_right
        '================================================================
        ' Function ajoutDroit
        ' Description : Ajout un droit pour un utilisateur
        '           Rend le droit ainsi crée ou nothing si Pb
        '================================================================
        Dim objRight As aut_right
        Dim oReturn As aut_right
        Try
            objRight = New aut_right(ptag, m_role, pDroit)
            objRight.text = pText
            m_colRights.Add(objRight, ptag)
            oReturn = objRight
        Catch ex As Exception
            setError("AjouteDroit", ex.ToString())
            oReturn = Nothing
        End Try
        Debug.Assert(Not oReturn Is Nothing, getErreur())
        Return oReturn
    End Function 'ajouteDroit

    Public Function AccesAuthorise(pstrMenuName As String) As Boolean
        Dim bReturn As Boolean
        Dim oRigth As aut_right
        Try
            bReturn = True
            'Si on trouve le nom du menu dans la collection , l'accès est interdit
            For Each oRigth In colRigths
                If oRigth.tag.ToUpper().Equals(pstrMenuName.ToUpper()) Then
                    bReturn = False
                    Exit For
                End If
            Next
        Catch ex As Exception
            setError("AccesAuthorise", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
#End Region

End Class
