Public Class Transporteur
    Inherits Tiers
    Private Shared m_colTransporteur As New Collection
    Private Shared m_bcolLoad As Boolean = False
    Private m_bdefaut As Boolean

    ''' <summary>
    ''' Rend une collection des transporteur.
    ''' Si bForce = True, la collection est rechargée.
    ''' </summary>
    ''' <param name="bForce"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property colTransporteur(Optional ByVal bForce As Boolean = False) As Collection
        Get
            If Not m_bcolLoad Or bForce Then
                shared_connect()
                m_colTransporteur = listeTRANSPORTEURS()
                m_bcolLoad = True
                shared_disconnect()
            End If
            Return m_colTransporteur
        End Get

    End Property
    Public Shared ReadOnly Property TransporteurDefault() As Transporteur
        Get
            Dim oCol As Collection
            Dim oTRP As Transporteur
            Dim oDefaut As Transporteur = Nothing
            oCol = colTransporteur
            For Each oTRP In oCol
                If oTRP.bDefaut Then
                    oDefaut = oTRP
                    Exit For
                End If
            Next
            Return oDefaut
        End Get
    End Property

    Public Sub New()
        MyBase.New("", "")
    End Sub

    Public Property bDefaut() As Boolean
        Get
            Return m_bdefaut
        End Get
        Set(ByVal Value As Boolean)
            m_bdefaut = Value
        End Set
    End Property


    Public Overrides Function checkForDelete() As Boolean
        Return False
    End Function

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        shared_connect()
        If pid <> 0 Then
            m_id = pid
        End If
        bReturn = loadTRP()
        shared_disconnect()
        Return bReturn

    End Function

    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        If bReturn Then
            bReturn = deleteTRP()
        End If
        Return bReturn

    End Function

    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        bReturn = InsertTRP()
        Return bReturn
    End Function

    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        bReturn = updateTRP()
        Return bReturn

    End Function

    Friend Overrides Function loadLight() As Boolean
        Return DBLoad()
    End Function

    Public Sub Dupplique(ByVal pTransp As Transporteur)
        m_bdefaut = pTransp.bDefaut
        m_id = pTransp.id
        code = pTransp.code            'code du tiers
        nom = pTransp.nom     'Nom du tiers
        rs = pTransp.rs        ' Raison Sociale
        AdresseLivraison.nom = pTransp.AdresseLivraison.nom
        AdresseLivraison.rue1 = pTransp.AdresseLivraison.rue1
        AdresseLivraison.rue2 = pTransp.AdresseLivraison.rue2
        AdresseLivraison.cp = pTransp.AdresseLivraison.cp
        AdresseLivraison.ville = pTransp.AdresseLivraison.ville
        AdresseLivraison.tel = pTransp.AdresseLivraison.tel
        AdresseLivraison.port = pTransp.AdresseLivraison.port
        AdresseLivraison.fax = pTransp.AdresseLivraison.fax
        AdresseLivraison.Email = pTransp.AdresseLivraison.Email
    End Sub

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objTRP As Tiers
        Dim bReturn As Boolean
        Try
            objTRP = CType(obj, Transporteur)
            '            bReturn = MyBase.Equals(obj)
            bReturn = True
            bReturn = bReturn And (nom.Equals(objTRP.nom))
            bReturn = bReturn And (id = objTRP.id)
            bReturn = bReturn And (code.Equals(objTRP.code))
            bReturn = bReturn And (rs.Equals(objTRP.rs))
            bReturn = bReturn And (Me.AdresseLivraison.nom.Equals(objTRP.AdresseLivraison.nom))
            bReturn = bReturn And (Me.AdresseLivraison.rue1.Equals(objTRP.AdresseLivraison.rue1))
            bReturn = bReturn And (Me.AdresseLivraison.rue2.Equals(objTRP.AdresseLivraison.rue2))
            bReturn = bReturn And (Me.AdresseLivraison.cp.Equals(objTRP.AdresseLivraison.cp))
            bReturn = bReturn And (Me.AdresseLivraison.ville.Equals(objTRP.AdresseLivraison.ville))
            bReturn = bReturn And (Me.AdresseLivraison.tel.Equals(objTRP.AdresseLivraison.tel))
            bReturn = bReturn And (Me.AdresseLivraison.fax.Equals(objTRP.AdresseLivraison.fax))
            bReturn = bReturn And (Me.AdresseLivraison.port.Equals(objTRP.AdresseLivraison.port))
            bReturn = bReturn And (Me.AdresseLivraison.Email.Equals(objTRP.AdresseLivraison.Email))


            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
End Class
