'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : lgPreComm    
' Description : Ligne de Precommande
'               une Ligne de précommande appartient à un Client
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
'===================================================================================================================================Public MustInherit Class Persist
Public Class lgPrecomm
    Inherits LgCommande

    Private m_idProduit As Integer
    Private m_codeProduit As String
    Private m_libProduit As String
    Private m_millesime As Integer
    Private m_libConditionnement As String
    Private m_QteHab As Decimal
    Private m_QteDern As Decimal
    Private m_PrixU As Double
    Private m_datedernCommande As Date
    Private m_refdernCommande As String


    Public Property idProduit() As Integer
        Get
            Return m_idProduit
        End Get
        Set(ByVal Value As Integer)
            m_idProduit = Value
        End Set
    End Property

    Public Property codeProduit() As String
        Get
            Return m_codeProduit
        End Get
        Set(ByVal Value As String)
            m_codeProduit = Value
        End Set
    End Property

    Public Property libProduit() As String
        Get
            Return m_libProduit
        End Get
        Set(ByVal Value As String)
            m_libProduit = Value
        End Set
    End Property
    Public Property millesime() As Integer
        Get
            Return m_millesime
        End Get
        Set(ByVal Value As Integer)
            m_millesime = Value
        End Set
    End Property
    Public Property libConditionnement() As String
        Get
            Return m_libConditionnement
        End Get
        Set(ByVal Value As String)
            m_libConditionnement = Value
        End Set
    End Property

    Public Property qteHab() As Decimal
        Get
            Return m_QteHab
        End Get
        Set(ByVal Value As Decimal)
            If m_QteHab <> Value Then
                m_QteHab = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property qteDern() As Decimal
        Get
            Return m_QteDern
        End Get
        Set(ByVal Value As Decimal)

            If m_QteDern <> Value Then
                m_QteDern = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Overloads Property prixU() As Double
        Get
            Return m_PrixU
        End Get
        Set(ByVal Value As Double)
            If (m_PrixU <> Value) Then
                m_PrixU = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property dateDerniereCommande() As Date
        Get
            Return m_datedernCommande.ToShortDateString
        End Get
        Set(ByVal Value As Date)
            If Not m_datedernCommande.ToShortDateString().Equals(Value.ToShortDateString()) Then
                m_datedernCommande = Value.ToShortDateString
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property refDerniereCommande() As String
        Get
            Return m_refdernCommande
        End Get
        Set(ByVal Value As String)
            If Not m_refdernCommande.Equals(Value) Then
                m_refdernCommande = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return m_idProduit & "," & m_codeProduit & "," & m_libProduit & "," & m_QteHab & "," & m_QteDern & "," & m_PrixU & "," & m_datedernCommande.ToShortDateString & "," & m_refdernCommande
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objLg As lgPrecomm
        Dim bReturn As Boolean
        Try
            bReturn = True
            objLg = CType(obj, lgPrecomm)

            '          bReturn = bReturn And (m_idProduit.Equals(objLg.id))
            'bReturn = bReturn And (idProduit.Equals(objLg.libProduit))
            'bReturn = bReturn And (m_millesime.Equals(objLg.millesime))
            'bReturn = bReturn And (m_libConditionnement.Equals(objLg.libConditionnement))
            bReturn = bReturn And (m_codeProduit.Equals(objLg.codeProduit))
            bReturn = bReturn And (m_QteHab.Equals(objLg.qteHab))
            bReturn = bReturn And (m_QteDern.Equals(objLg.qteDern))
            bReturn = bReturn And (m_PrixU.Equals(objLg.prixU))
            bReturn = bReturn And (m_datedernCommande.Equals(objLg.dateDerniereCommande))
            bReturn = bReturn And (m_refdernCommande.Equals(objLg.refDerniereCommande))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Friend Sub New(ByVal IdPrecommande As Integer, Optional ByVal idProduit As Integer = 0, Optional ByVal strCodeProduit As String = "", Optional ByVal strLibProduit As String = "", Optional ByVal dQteHab As Decimal = 0, Optional ByVal dQteDern As Decimal = 0, Optional ByVal pPrixU As Double = 0, Optional ByVal pDateDern As Date = DATE_DEFAUT, Optional ByVal prefdernCommande As String = "")
        MyBase.New(IdPrecommande)
        num = idProduit
        m_idProduit = idProduit
        m_codeProduit = strCodeProduit
        m_libProduit = strLibProduit
        m_QteHab = dQteHab
        m_QteDern = dQteDern
        m_PrixU = pPrixU
        m_datedernCommande = pDateDern
        m_refdernCommande = prefdernCommande
    End Sub 'New

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_codeProduit
        End Get
    End Property


    Public Overrides Function checkForDelete() As Boolean
        Return True
    End Function

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Return False
    End Function

    Friend Overrides Function delete() As Boolean
        Return False
    End Function

    Friend Overrides Function insert() As Boolean
        Return False
    End Function

    Friend Overrides Function update() As Boolean
        Return False
    End Function
End Class
