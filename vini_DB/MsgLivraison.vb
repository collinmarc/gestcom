Public Class MsgLivraison
    Private m_id As Long
    Public Property Id() As Long
        Get
            Return m_id
        End Get
        Set(ByVal value As Long)
            m_id = value
        End Set
    End Property
    Private m_numCommande As String
    Public Property NumCommande() As String
        Get
            Return m_numCommande
        End Get
        Set(ByVal value As String)
            m_numCommande = value
        End Set
    End Property
    Private m_dateMsg As Date
    Public Property DateMessage() As Date
        Get
            Return m_dateMsg
        End Get
        Set(ByVal value As Date)
            m_dateMsg = value
        End Set
    End Property
    Private m_nbreColisLivre As Integer
    Public Property NbreColisLivre() As Integer
        Get
            Return m_nbreColisLivre
        End Get
        Set(ByVal value As Integer)
            m_nbreColisLivre = value
        End Set

    End Property
    Private m_Message As String
    Public Property Message() As String
        Get
            Return m_Message
        End Get
        Set(ByVal value As String)
            m_Message = value
        End Set
    End Property
    Private m_bMsgTraite As Boolean
    Public Property MsgTraite() As Boolean
        Get
            Return m_bMsgTraite
        End Get
        Set(ByVal value As Boolean)
            m_bMsgTraite = value
        End Set
    End Property
    ''' <summary>
    ''' Resulat du traitement
    ''' 0=OK
    ''' 1=ERREUR DE LIVRAISON
    ''' 2=ERREUR DE TRAITEMENT
    ''' </summary>
    ''' <remarks></remarks>
    Private m_Resultat As Integer
    Public Property Resultat() As Integer
        Get
            Return m_Resultat
        End Get
        Set(ByVal value As Integer)
            m_Resultat = value
        End Set
    End Property


End Class
