'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : lgFactColisage
' Description : Ligne de Facture de Colisage
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
Public Class LgFactColisage
    Inherits Persist

    Private m_num As Integer                'Numéro d'ordre de ligne
    Private m_idFactColisage As Long
    Private m_dDeb As Date
    Private m_dFin As Date
    Private m_StockInitial As Integer
    Private m_Entrees As Integer
    Private m_Sorties As Integer
    Private m_StockFinal As Integer
    Private m_qte As Integer
    Private m_pu As Decimal
    Private m_Montant_HT As Decimal



#Region "Accesseurs"
    Public Property num() As Integer
        Get
            Return m_num
        End Get
        Set(ByVal Value As Integer)
            If m_num <> Value Then
                m_num = Value
                bUpdated = True
            End If
        End Set
    End Property
    Public Sub New()
        m_typedonnee = vncEnums.vncTypeDonnee.LGFACTCOL
        m_idFactColisage = 0
        m_dDeb = DATE_DEFAUT
        m_dFin = DATE_DEFAUT
        m_StockInitial = 0
        m_Entrees = 0
        m_Sorties = 0
        m_StockFinal = 0
        m_qte = 0
        m_pu = 0.0
        m_Montant_HT = 0.0
    End Sub
    Public Property idFactColisage() As Long
        Get
            Return m_idFactColisage
        End Get
        Set(ByVal Value As Long)
            If Value <> m_idFactColisage Then
                m_idFactColisage = Value
                RaiseUpdated()
            End If
        End Set
    End Property


    Public Property dDeb() As Date
        Get
            Return m_dDeb
        End Get
        Set(ByVal Value As Date)
            If m_dDeb <> Value Then
                m_dDeb = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property dFin() As Date
        Get
            Return m_dFin
        End Get
        Set(ByVal Value As Date)
            If m_dFin <> Value Then
                m_dFin = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property StockInitial() As Integer
        Get
            Return m_StockInitial
        End Get
        Set(ByVal Value As Integer)
            If (Not Value.Equals(m_StockInitial)) Then
                m_StockInitial = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property StockFinal() As Integer
        Get
            Return m_StockFinal
        End Get
        Set(ByVal Value As Integer)
            If (Not Value.Equals(m_StockFinal)) Then
                m_StockFinal = Value
                m_qte = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Entrees() As Integer
        Get
            Return m_Entrees
        End Get
        Set(ByVal Value As Integer)
            If (Not Value.Equals(m_Entrees)) Then
                m_Entrees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property Sorties() As Integer
        Get
            Return m_Sorties
        End Get
        Set(ByVal Value As Integer)
            If (Not Value.Equals(m_Sorties)) Then
                m_Sorties = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qte() As Integer
        Get
            Return m_qte
        End Get
        Set(ByVal Value As Integer)
            If (Not Value.Equals(m_qte)) Then
                m_qte = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property pu() As Decimal
        Get
            Return m_pu
        End Get
        Set(ByVal Value As Decimal)
            If m_pu <> Value Then
                m_pu = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property MontantHT() As Decimal
        Get
            Return m_Montant_HT
        End Get
        Set(ByVal Value As Decimal)
            If m_Montant_HT <> Value Then
                m_Montant_HT = Value
                RaiseUpdated()
            End If
        End Set
    End Property

#End Region
#Region "Méthodes Persist"

    '=======================================================================
    'Fonction : DBLoad()
    'Description : Chargement de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Return bReturn
    End Function 'DBLoad

    '=======================================================================
    'Fonction : delete()
    'Description : Suppression de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai si l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Return bReturn

    End Function ' delete
    '=======================================================================
    'Fonction : checkFordelete
    'description : Controle si l'élément est supprimable
    '                table commandesClients
    '=======================================================================
    Public Overrides Function checkForDelete() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete

    '=======================================================================
    'Fonction : insert()
    'Description : insertion de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean

        bReturn = False
        Return bReturn

    End Function 'insert

    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Return bReturn

    End Function 'Update
#End Region

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "LGFACTCOLISAGE =(" & m_id & "," & m_idFactColisage & "," & m_Montant_HT & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objLgFactColisage As LgFactColisage
        Try
            objLgFactColisage = CType(obj, LgFactColisage)
            bReturn = True
            bReturn = bReturn And (m_num.Equals(objLgFactColisage.num))
            bReturn = bReturn And (m_idFactColisage = objLgFactColisage.idFactColisage)
            bReturn = bReturn And (m_Montant_HT = objLgFactColisage.MontantHT)
            bReturn = bReturn And (m_dDeb.Equals(objLgFactColisage.dDeb))
            bReturn = bReturn And (m_dFin.Equals(objLgFactColisage.dFin))
            bReturn = bReturn And (m_StockInitial.Equals(objLgFactColisage.StockInitial))
            bReturn = bReturn And (m_StockFinal.Equals(objLgFactColisage.StockFinal))
            bReturn = bReturn And (m_Entrees = objLgFactColisage.Entrees)
            bReturn = bReturn And (m_Sorties = objLgFactColisage.Sorties)
            bReturn = bReturn And (m_qte = objLgFactColisage.qte)
            bReturn = bReturn And (m_pu = objLgFactColisage.pu)
            bReturn = bReturn And (m_Montant_HT = objLgFactColisage.MontantHT)
        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'Equals

    '=======================================================================
    'Fonction : calculPrixTotal()
    'Description : Calcul du prix total
    'Détails    :  
    'Retour : 
    '=======================================================================
    Public Sub calculPrixTotal()
        m_Montant_HT = m_qte * pu
    End Sub
    Public Sub calculStockFinal()
        m_StockFinal = m_StockInitial + m_Entrees + m_Sorties
        m_qte = m_StockFinal
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "LGCOLISAGE"
        End Get
    End Property

End Class
