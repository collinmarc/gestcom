'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : <<LgStat>>
' Description : <<Ligne de statistiques >>
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
Public Class LgStat
    Inherits Persist

    Private m_cmd_code As String            'Code de la Commande
    Private m_cmd_date As Date              'Date de la commande
    Private m_cmd_total_HT As Decimal       'Total HT de la commande
    Private m_cmd_total_TTC As Decimal      'Total TTC de la Commande
    Private m_cmd_clt_id As Integer         'Id du client
    Private m_lg_prd_id As Integer         'Id du produit
    Private m_lg_frn_id As Integer          'IF du fournisseur du produit
    Private m_lg_qte_commandee As Decimal   'Qte commandée
    Private m_lg_qte_livree As Decimal   'Qte commandée
    Private m_lg_qte_facturee As Decimal   'Qte commandée
    Private m_lg_prixU As Decimal           'Prix Unitaire
    Private m_lg_total_HT As Decimal        'Montant HT de la ligne
    Private m_lg_total_TTC As Decimal     'Montant TTC de la ligne

#Region "Accesseurs"
    Public Sub New(ByVal objCMD As CommandeClient, ByVal objLg As LgCommande)
        '===================================================================
        ' Function :contructeur
        'description : contruction d'une ligne de stat à partir des commandes et des lignes
        '===================================================================
        Debug.Assert(Not objCMD Is Nothing)
        Debug.Assert(Not objLg Is Nothing)

        cmd_code = objCMD.code
        cmd_date = objCMD.dateCommande
        cmd_clt_id = objCMD.oTiers.id
        cmd_total_HT = objCMD.totalHT
        cmd_total_TTC = objCMD.totalTTC
        lg_prd_id = objLg.oProduit.id
        lg_frn_id = objLg.oProduit.idFournisseur
        lg_qte_commandee = objLg.qteCommande
        lg_qte_livree = objLg.qteLiv
        lg_qte_facturee = objLg.qteFact
        lg_prixU = objLg.prixU
        lg_total_ht = objLg.prixHT
        lg_total_TTC = objLg.prixTTC

    End Sub


    Public Property cmd_code() As String
        Get
            Return m_cmd_code
        End Get
        Set(ByVal Value As String)
            m_cmd_code = Value
        End Set
    End Property

    Public Property cmd_date() As Date
        Get
            Return m_cmd_date.ToShortDateString
        End Get
        Set(ByVal Value As Date)
            m_cmd_date = Value.ToShortDateString
        End Set
    End Property

    Public Property cmd_total_HT() As Decimal
        Get
            Return m_cmd_total_HT
        End Get
        Set(ByVal Value As Decimal)
            m_cmd_total_HT = Value
        End Set
    End Property

    Public Property cmd_total_TTC() As Decimal
        Get
            Return m_cmd_total_TTC
        End Get
        Set(ByVal Value As Decimal)
            m_cmd_total_TTC = Value
        End Set
    End Property
    Public Property cmd_clt_id() As Integer
        Get
            Return m_cmd_clt_id
        End Get
        Set(ByVal Value As Integer)
            m_cmd_clt_id = Value
        End Set
    End Property

    Public Property lg_prd_id() As Integer
        Get
            Return m_lg_prd_id
        End Get
        Set(ByVal Value As Integer)
            m_lg_prd_id = Value
        End Set
    End Property

    Public Property lg_frn_id() As Integer
        Get
            Return m_lg_frn_id
        End Get
        Set(ByVal Value As Integer)
            m_lg_frn_id = Value

        End Set
    End Property

    Public Property lg_qte_commandee() As Decimal
        Get
            Return m_lg_qte_commandee
        End Get
        Set(ByVal Value As Decimal)
            m_lg_qte_commandee = Value

        End Set
    End Property

    Public Property lg_qte_livree() As Decimal
        Get
            Return m_lg_qte_livree
        End Get
        Set(ByVal Value As Decimal)
            m_lg_qte_livree = Value

        End Set
    End Property
    Public Property lg_qte_facturee() As Decimal
        Get
            Return m_lg_qte_facturee
        End Get
        Set(ByVal Value As Decimal)
            m_lg_qte_facturee = Value

        End Set
    End Property
    Public Property lg_prixU() As Decimal
        Get
            Return m_lg_prixU
        End Get
        Set(ByVal Value As Decimal)
            m_lg_prixU = Value

        End Set
    End Property

    Public Property lg_total_ht() As Decimal
        Get
            Return m_lg_total_HT
        End Get
        Set(ByVal Value As Decimal)
            m_lg_total_HT = Value
        End Set
    End Property
    Public Property lg_total_TTC() As Decimal
        Get
            Return m_lg_total_TTC
        End Get
        Set(ByVal Value As Decimal)
            m_lg_total_TTC = Value
        End Set
    End Property
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "LGSTAT =(" & ")"
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
#End Region

#Region "Interface Persist"

    Public Overrides Function checkForDelete() As Boolean

    End Function

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean

    End Function

    Friend Overrides Function delete() As Boolean

    End Function

    Friend Overrides Function insert() As Boolean
        Return False
    End Function

    Friend Overrides Function update() As Boolean

    End Function
#End Region
End Class
