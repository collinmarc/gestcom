''' <summary>
''' Classe Ligne de facture Hobivin
''' Créée à partir de ligneCommande
''' </summary>
''' <remarks></remarks>
Public Class LgFactHBV
    Inherits Persist

    Private m_num As Integer                'Numéro d'ordre de ligne
    Private m_oProduit As Produit           'Produit Commandé
    Private m_idFactHBV As Integer                 'Id de la Facture
    Private m_qteCom As Decimal         'Quantite commandée
    Private m_qteLiv As Decimal         'Quantite commandée
    Private m_qteFact As Decimal         'Quantite commandée
    Private m_prixU As Decimal               'Prix Unitaire
    Private m_prixHT As Decimal               'Prix Total HT
    Private m_prixTTC As Decimal               'Prix Total TTC
    Private m_bGratuit As Boolean           'Ligne Gratuite
    Private m_bProduitLoaded As Boolean

#Region "Accesseurs"
    Public Sub New()
        m_typedonnee = vncEnums.vncTypeDonnee.LGFACTHBV
        m_idFactHBV = 0
        m_num = 0
        m_bGratuit = False
        m_oProduit = Nothing
        m_prixHT = 0
        m_prixTTC = 0
        m_prixU = 0
        m_qteCom = 0
        m_qteLiv = 0
        m_qteFact = 0
        m_bProduitLoaded = False
    End Sub 'New

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

    Public Property oProduit() As Produit
        Get
            Return m_oProduit
        End Get
        Set(ByVal Value As Produit)
            If m_oProduit Is Nothing Then
                If Not Value Is Nothing Then
                    m_oProduit = Value
                    RaiseUpdated()
                End If
            Else
                If Not m_oProduit.Equals(Value) Then
                    m_oProduit = Value
                    RaiseUpdated()
                End If
            End If

        End Set
    End Property
    Public ReadOnly Property ProduitCode() As String
        Get
            Return oProduit.code
        End Get
    End Property
    Public ReadOnly Property ProduitNom() As String
        Get
            Return oProduit.nom
        End Get
    End Property
    Public ReadOnly Property ProduitMil() As String
        Get
            Return oProduit.millesime
        End Get
    End Property
    Public ReadOnly Property ProduitConditionnement() As String
        Get
            Return oProduit.libConditionnement
        End Get
    End Property
    Public ReadOnly Property ProduitContenant() As String
        Get
            Return oProduit.libContenant
        End Get
    End Property
    Public ReadOnly Property ProduitCouleur() As String
        Get
            Return oProduit.libCouleur
        End Get
    End Property
    Public Property idFactHBV() As Integer
        Get
            Return m_idFactHBV
        End Get
        Set(ByVal Value As Integer)
            If m_idFactHBV <> Value Then
                m_idFactHBV = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qteCommande() As Decimal
        Get
            Return m_qteCom
        End Get
        Set(ByVal Value As Decimal)
            If m_qteCom <> Value Then
                m_qteCom = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    Public Property qteLiv() As Decimal
        Get
            Return m_qteLiv
        End Get
        Set(ByVal Value As Decimal)
            If m_qteLiv <> Value Then
                m_qteLiv = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    Public Property qteFact() As Decimal
        Get
            Return m_qteFact
        End Get
        Set(ByVal Value As Decimal)
            If m_qteFact <> Value Then
                m_qteFact = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property prixU() As Decimal
        Get
            Return m_prixU
        End Get
        Set(ByVal Value As Decimal)
            If m_prixU <> Value Then
                m_prixU = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    Public Property prixHT() As Decimal
        Get
            Return m_prixHT
        End Get
        Set(ByVal Value As Decimal)
            If m_prixHT <> Value Then
                m_prixHT = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property prixTTC() As Decimal
        Get
            Return m_prixTTC
        End Get
        Set(ByVal Value As Decimal)
            If m_prixTTC <> Value Then
                m_prixTTC = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property bGratuit() As Boolean
        Get
            Return m_bGratuit
        End Get
        Set(ByVal Value As Boolean)
            If m_bGratuit <> Value Then
                m_bGratuit = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return oProduit.code & " " & m_qteCom
        End Get
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
        Return "LGFHBV =(" & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objLgCommande As LgCommande
        Try

            bReturn = obj.GetType.Name.Equals(Me.GetType().Name)
            objLgCommande = obj
            bReturn = MyBase.Equals(obj)
            bReturn = bReturn And (num = (objLgCommande.num))
            bReturn = bReturn And (idFactHBV = (objLgCommande.idCmd))
            If Not oProduit Is Nothing Then
                bReturn = bReturn And (oProduit.Equals(objLgCommande.oProduit))
            Else
                bReturn = objLgCommande.oProduit Is Nothing
            End If
            bReturn = bReturn And (qteCommande = objLgCommande.qteCommande)
            bReturn = bReturn And (qteLiv = objLgCommande.qteLiv)
            bReturn = bReturn And (qteFact = objLgCommande.qteFact)
            bReturn = bReturn And (prixU = objLgCommande.prixU)
            bReturn = bReturn And (prixHT = objLgCommande.prixHT)
            bReturn = bReturn And (prixTTC = objLgCommande.prixTTC)
            bReturn = bReturn And (bGratuit = objLgCommande.bGratuit)


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
    Public Function calculPrixTotal() As Boolean
        Dim oParam As Param
        Dim nTaux As Decimal
        Dim qte As Integer

        Debug.Assert(Not oProduit Is Nothing, "Produit Non renseigné")
        If m_bGratuit Then
            m_prixU = 0
            m_prixHT = 0
            m_prixTTC = 0
        Else
            'Récupération du taux de TVA du produit
            For Each oParam In Param.colTVA
                If oParam.id = oProduit.idTVA Then
                    nTaux = CDbl(oParam.valeur)
                    Exit For
                End If
            Next
            Debug.Assert(nTaux <> 0)

            qte = qteFact
            m_prixHT = m_prixU * qte
            m_prixTTC = Math.Round(m_prixHT * (1 + (nTaux / 100)), 2)
        End If
        Return True
    End Function 'calculPrixTotal
End Class
