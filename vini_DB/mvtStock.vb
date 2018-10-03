Imports System.Collections.Generic
''' <summary>
''' Classe mouvement de Stock
''' </summary>
''' <remarks></remarks>
Public Class mvtStock
    Inherits Persist

    Private m_idProduit As Integer              'Produit 
    Private m_date As Date                      'Date du mouvement
    Private m_TypeMvt As vncTypeMvt             'Type de Mouvement
    Private m_idReference As Integer            'id de la référence
    Private m_libelle As String                 'Libellé du mouvement
    Private m_qte As Decimal                       'Quantite 
    Private m_oCommentaire As Commentaire       'Commentaire
    Private m_bControle As Boolean              ' mvt controle
    Private m_Etat As EtatMvtStock              ' Etat du Mvt de Stock
    Private m_idFactColisage As Integer              ' Id de la facture de colisage

#Region "Accesseurs"
    Public Sub New(ByVal pdateMvt As Date, ByVal pidProduit As Integer, ByVal ptype As vncTypeMvt, ByVal pqte As Decimal, ByVal plib As String)
        m_typedonnee = vncEnums.vncTypeDonnee.MVTSTK
        m_idProduit = pidProduit
        m_date = pdateMvt.ToShortDateString
        m_TypeMvt = ptype
        m_idReference = 0
        m_libelle = plib
        m_qte = pqte
        m_oCommentaire = New Commentaire
        m_bControle = False
        m_Etat = EtatMvtStock.createEtat(vncEtatMVTSTK.vncMVTSTK_nFact)
        m_idFactColisage = 0
        majBooleenAlaFinDuNew()
    End Sub 'New
    Friend Sub New()
        m_typedonnee = vncEnums.vncTypeDonnee.MVTSTK
        m_idProduit = 0
        m_date = DATE_DEFAUT
        m_TypeMvt = vncEnums.vncTypeMvt.vncmvtRegul
        m_idReference = 0
        m_libelle = ""
        m_qte = 0
        m_oCommentaire = New Commentaire
        m_bControle = False
        m_Etat = EtatMvtStock.createEtat(vncEtatMVTSTK.vncMVTSTK_nFact)
        m_idFactColisage = 0
    End Sub 'New

    Public Property idProduit() As Integer
        Get
            Return m_idProduit
        End Get
        Set(ByVal Value As Integer)
            If Not m_idProduit.Equals(Value) Then
                m_idProduit = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property datemvt() As Date
        Get
            Return m_date
        End Get
        Set(ByVal Value As Date)
            If m_date.ToShortDateString <> Value.ToShortDateString Then
                m_date = Value.ToShortDateString
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property typeMvt() As vncTypeMvt
        Get
            Return m_TypeMvt
        End Get
        Set(ByVal Value As vncTypeMvt)
            If Value <> m_TypeMvt Then
                RaiseUpdated()
                m_TypeMvt = Value
            End If
        End Set
    End Property
    Public Property typeMvtSTR() As String
        Get
            Return typeMvt
        End Get
        Set(ByVal Value As String)
            If Not m_TypeMvt.Equals(Value) Then
                Try
                    typeMvt = Value
                    RaiseUpdated()
                Catch
                End Try

            End If
        End Set
    End Property
    Public Property idReference() As Integer
        Get
            Return m_idReference
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idReference Then
                RaiseUpdated()
                m_idReference = Value
            End If
        End Set
    End Property
    Public Property libelle() As String
        Get
            Return m_libelle
        End Get
        Set(ByVal Value As String)
            If Value <> m_libelle Then
                RaiseUpdated()
                m_libelle = Value
            End If
        End Set
    End Property
    Public Property qte() As Decimal
        Get
            Return m_qte
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_qte Then
                RaiseUpdated()
                m_qte = Value
            End If
        End Set
    End Property
    Public Property Commentaire() As String
        Get
            Return m_oCommentaire.comment
        End Get
        Set(ByVal Value As String)
            If Not (Value.Equals(m_oCommentaire.comment)) Then
                RaiseUpdated()
                m_oCommentaire.comment = Value
            End If
        End Set
    End Property

    Public Property Etat() As EtatMvtStock
        Get
            Return m_Etat
        End Get
        Set(ByVal value As EtatMvtStock)
            If Not (value.Equals(m_Etat)) Then
                RaiseUpdated()
                m_Etat = value
            End If

        End Set
    End Property

    Public Property idFactColisage() As Integer
        Get
            Return m_idFactColisage
        End Get
        Set(ByVal value As Integer)
            If Not (value.Equals(m_idFactColisage)) Then
                RaiseUpdated()
                m_idFactColisage = value
            End If
        End Set
    End Property
    Public ReadOnly Property key() As String
        'on veut trier les mvts de stocks dans l'odre desc des date et des type de Mvts
        ' Donc on fait une clé sur une différence avec le 3eme millénaire
        ' et la différence 99-
        Get
            Dim nDate As Long
            nDate = 30000101 - ((m_date.Year * 10000) + (m_date.Month * 100) + m_date.Day)
            Return Format(nDate, "0000000000") & "_" & (99 - m_TypeMvt)
        End Get
    End Property

    Public Property bControle() As Boolean
        Get
            Return m_bControle
        End Get
        Set(ByVal Value As Boolean)
            m_bControle = Value
        End Set
    End Property
#End Region
#Region "Interface Persist"
    '=======================================================================
    'Fonction : DBLoad()
    'Description : Chargement de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        Debug.Assert(id <> 0, "idCommande <> 0")
        bReturn = loadMVTSTK()
        Return bReturn
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : delete()
    'Description : Suppression de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai si l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        bReturn = deleteMVTSTK()
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
        Debug.Assert(m_idProduit <> 0, "Le Produit n'est pas Renseigné")
        Debug.Assert(id = 0, "idmvt = 0")

        Dim bReturn As Boolean
        bReturn = insertMVTSTK()
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Debug.Assert(m_idProduit <> 0, "Le Produit n'est pas Renseigné")
        Debug.Assert(id <> 0, "id <> 0")
        Dim bReturn As Boolean
        bReturn = updateMVTSTK()
        Return bReturn

    End Function 'Update

    Public Shared Function getListeColisage(ByVal pidFactColisage As Integer) As ColEventSorted
        '=======================================================================
        'Fonction : getListe()
        'Description : Rend une liste des mvt de stocks pour un facture de colisage
        'Détails    :  Si ID Commande est renseigné alors on ajoute un filtre sur le type de Mvt = "2"
        '               Si ID BA est renseigné alors on ajoute un filtre sur le type de Mvt = "3"
        '               Sinon on retourne tous les mvts de stocks
        'Retour : collection des mouvements de stock
        '=======================================================================
        Dim colReturn As ColEventSorted

        Persist.shared_connect()
        colReturn = Persist.ListeMVTSTK_FACTCOL(pidFactColisage)
        Persist.shared_disconnect()
        Return colReturn
    End Function 'getListeColisage
    Public Shared Function getListe(ByVal pidProduit As Integer, Optional ByVal pidCmd As Integer = -1, Optional ByVal pidBA As Integer = -1) As ColEventSorted
        '=======================================================================
        'Fonction : getListe()
        'Description : Rend une liste des mvt de stocks
        'Détails    :  Si ID Commande est renseigné alors on ajoute un filtre sur le type de Mvt = "2"
        '               Si ID BA est renseigné alors on ajoute un filtre sur le type de Mvt = "3"
        '               Sinon on retourne tous les mvts de stocks
        'Retour : collection des mouvements de stock
        '=======================================================================
        Dim colReturn As ColEventSorted

        Persist.shared_connect()
        If pidCmd <> -1 Then
            colReturn = Persist.ListeMVTSTK(pidProduit, vncEnums.vncTypeMvt.vncMvtCommandeClient, pidCmd)
        Else
            If pidBA <> -1 Then
                colReturn = Persist.ListeMVTSTK(pidProduit, vncEnums.vncTypeMvt.vncmvtBonAppro, pidBA)
            Else
                colReturn = Persist.ListeMVTSTK(pidProduit)
            End If
        End If
        Persist.shared_disconnect()
        Return colReturn
    End Function 'getListe
    ''' <summary>
    ''' Chergement de la liste des Mvt de stocks depuis le dernier inventaire
    ''' </summary>
    ''' <param name="pidProduit">ID du produit</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getListeDepuisDernMvtInventaire(ByVal pidProduit As Integer) As List(Of mvtStock)
        '=======================================================================
        'Fonction : getListe()
        'Description : Rend une liste des mvt de stocks
        'Détails    :  Si ID Commande est renseigné alors on ajoute un filtre sur le type de Mvt = "2"
        '               Si ID BA est renseigné alors on ajoute un filtre sur le type de Mvt = "3"
        '               Sinon on retourne tous les mvts de stocks
        'Retour : collection des mouvements de stock
        '=======================================================================
        Dim colReturn As New List(Of mvtStock)

        Persist.shared_connect()
        colReturn = Persist.ListeMVTSTKDepuisDernMvtInventaire(pidProduit)
        Persist.shared_disconnect()
        Return colReturn
    End Function 'getListe

    Public Shared Function getListe2(ByVal pDateDebut As Date, ByVal pDateFin As Date, Optional ByVal pFournisseur As Fournisseur = Nothing, Optional ByVal pEtat As vncEtatMVTSTK = vncEtatMVTSTK.vncMVTSTK_Tous) As Collection
        '=======================================================================
        'Fonction : getListe()
        'Description : Rend une liste des mvt de stocks pour les produits en Stock avec un filtre éventuel sur l'ID du fournisseur
        '           les stocks sont classé par ordre chronologique !!!
        'Retour : collection des mouvements de stock
        '=======================================================================
        Dim colReturn As Collection

        Persist.shared_connect()
        colReturn = Persist.ListeMVTSTK2(pDateDebut, pDateFin, pFournisseur, pEtat)
        Persist.shared_disconnect()
        Return colReturn
    End Function 'getListe
#End Region

    '=======================================================================
    'Fonction : shortResume()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_date.ToShortDateString & "|" & m_qte & "|" & m_libelle
        End Get
    End Property

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "MVT =(" & m_TypeMvt & "," & m_idProduit & "," & m_date & "," & m_qte & "," & m_libelle & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objMvt As mvtStock
        Try

            bReturn = obj.GetType.Name.Equals(Me.GetType().Name)
            objMvt = obj
            bReturn = MyBase.Equals(obj)
            bReturn = bReturn And (m_idProduit.Equals(objMvt.idProduit))
            bReturn = bReturn And (m_date.ToShortDateString.Equals(objMvt.datemvt.ToShortDateString))
            bReturn = bReturn And (m_TypeMvt.Equals(objMvt.typeMvt))
            bReturn = bReturn And (m_idReference.Equals(objMvt.idReference))
            bReturn = bReturn And (m_libelle.Equals(objMvt.libelle))
            bReturn = bReturn And (m_qte.Equals(objMvt.qte))
            bReturn = bReturn And (m_oCommentaire.comment.Equals(objMvt.Commentaire))
            bReturn = bReturn And (m_Etat.Equals(objMvt.Etat))
            bReturn = bReturn And (m_idFactColisage.Equals(objMvt.idFactColisage))

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'Equals

    Public Sub changeEtat(ByVal pEtat As vncActionFactColisage)
        m_Etat = m_Etat.changeEtat(pEtat)
    End Sub



End Class
