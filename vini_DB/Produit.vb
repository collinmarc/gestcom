Imports System.Collections.Generic
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Produit
' Description : Un produit de l'application
'===================================================================================================================================
'Heritage
'=========
'       Persist
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
' 17/10/05 : Ajout du poids calculé
'
'===================================================================================================================================Public MustInherit Class Persist
Public Class Produit
    Inherits Persist
    '==============================================================================
    Private m_code As String                'Code Article
    Private m_nom As String                 'Designation de l'article
    Private m_motcle As String              'Mot clé de recherche
    Private m_millesime As Integer          'Millésime
    Private m_idCouleur As Integer 'Id de la couleur
    Private m_idContenant As Integer           'Id du contenant
    Private m_libContenant As String        'Libellé du contenant
    Private m_idConditionnement As Integer     ' Id du conditionnement
    Private m_idRegion As Integer              'id Region vinicole
    Private m_bDisponible As Boolean        'Produit au catalogue
    Private m_idTVA As Integer                 'Id de TVA
    Private m_codeStat As String            'Code Statistique
    Private m_QteStock As Decimal             'Quantité en Stock
    Private m_DateDernInventaire As Date    'Date de Dernier inventaire
    Private m_QteStockDernInventaire As Decimal 'Qte en stock au dernier inventaire
    Private m_idFournisseur As Integer         'Id du Fournisseur
    Private m_nomFournisseur As String      'Nom du Fournisseur
    Private m_qteCommande As Decimal           'qte en commande
    Private m_TarifA As Decimal           'Tarif 
    Private m_TarifB As Decimal           'Tarif
    Private m_TarifC As Decimal           'Tarif
    Private m_bStock As Boolean             'Produit Stocké
    Private WithEvents m_colMvtStock As ColEventSorted       'Collection des Mouvements de Stocks
    Private m_bcolMvtStockLoaded As Boolean
    Private m_bcolMvtStockUpdated As Boolean
    Private m_Dossier As String

    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des Produits
    'Retour : Rend une collection de fournisseurs
    '=======================================================================
    Public Shared Function getListe(ByVal pTypeProduit As vncTypeProduit, Optional ByVal strCode As String = "", Optional ByVal strlibelle As String = "", Optional ByVal strMotCle As String = "", Optional ByVal idFournisseur As Integer = 0, Optional ByVal idClient As Integer = 0, Optional pdossier As String = "") As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = ListePRD(pTypeProduit, strCode, strlibelle, strMotCle, idFournisseur, idClient, pdossier)
        shared_disconnect()
        Return colReturn
    End Function
    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : recalculdesStocks 
    'Description : Recalcul des stocks de tous les produits
    'Retour : Vrai ou Faux
    '=======================================================================
    Public Shared Function recalculdesStocks() As Boolean
        Dim objProduit As Produit
        Dim colProduit As Collection
        Dim bReturn As Boolean
        Try
            'Chargement de tous les produits plateforme
            colProduit = Produit.getListe(vncTypeProduit.vncPlateforme, )

            bReturn = True
            'Pour Chaque produit
            For Each objProduit In colProduit
                'Chargement du produit
                bReturn = objProduit.load()
                Debug.Assert(bReturn, "objProduit.load" & Produit.getErreur())
                bReturn = objProduit.loadcolmvtStock()
                Debug.Assert(bReturn, "objProduit.loadcolmvtStock" & Produit.getErreur())
                'Recalcul de son stock
                bReturn = objProduit.recalculStock()
                If bReturn Then
                    bReturn = objProduit.save()
                End If
            Next
        Catch ex As Exception
            bReturn = False
            setError("Produit.recalculdesStocks", ex.ToString)
        End Try

        Debug.Assert(bReturn, "objProduit.reCalculStock" & Produit.getErreur())
        Return bReturn
    End Function 'recalculdesStocks

#Region "Accesseurs"
    Friend Sub New()
        init()
        m_bcolMvtStockUpdated = False
        Debug.Assert(m_bNew, "bNew")
    End Sub 'New
    Private Sub init()
        m_typedonnee = vncEnums.vncTypeDonnee.PRODUIT
        m_code = ""
        m_nom = ""
        m_motcle = ""
        m_bDisponible = True
        m_bStock = True
        m_idTVA = Param.TVAdefaut.id
        m_codeStat = ""
        m_QteStock = 0
        m_DateDernInventaire = DATE_DEFAUT
        m_QteStockDernInventaire = 0
        m_idFournisseur = 0
        m_nomFournisseur = ""
        m_colMvtStock = New ColEventSorted
        m_TarifA = 0
        m_TarifB = 0
        m_TarifC = 0
        m_Dossier = vini_DB.Dossier.VINICOM
        'Pour un nouveau produit on considère que la collection est chargée,
        'elle sera considérée comme non-chargée lors du load du produit
        m_bcolMvtStockLoaded = True
        m_bcolMvtStockUpdated = False
        setQteCommande(0)
    End Sub
    Public Sub New(ByVal strCode As String, ByVal oFRN As Fournisseur, ByVal pMil As Integer)
        Debug.Assert(Param.couleurdefaut.defaut, "Pas de Couleur par defaut")
        Debug.Assert(contenant.contenantDefaut.defaut, "Pas de contenant par defaut")
        Debug.Assert(Param.conditionnementdefaut.defaut, "Pas de Conditionnement par defaut")
        Debug.Assert(Param.regiondefaut.defaut, "Pas de Region par defaut")
        Debug.Assert(Param.TVAdefaut.defaut, "Pas de TVA par defaut")
        init()
        m_code = strCode
        m_nom = ""
        m_motcle = ""
        m_millesime = 1990
        m_idCouleur = Param.couleurdefaut.id
        m_idContenant = contenant.contenantDefaut.id
        m_libContenant = contenant.contenantDefaut.libelle
        m_idConditionnement = Param.conditionnementdefaut.id
        m_idRegion = Param.regiondefaut.id
        m_bDisponible = True
        m_bStock = True
        m_idTVA = Param.TVAdefaut.id
        m_codeStat = strCode
        If oFRN Is Nothing Then
            m_idFournisseur = 0
            m_nomFournisseur = ""
        Else
            m_idFournisseur = oFRN.id
            m_nomFournisseur = oFRN.nom
        End If
        m_colMvtStock = New ColEventSorted
        'Pour un nouveau produit on considère que la collection est chargée,
        'elle sera considérée comme non-chargée lors du load du produit
        m_bcolMvtStockLoaded = True
        Debug.Assert(m_bNew, "bNew")

    End Sub 'New
    Public Shared Function createandload(ByVal pid As Integer) As Produit
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim objProduit As Produit
        Dim bReturn As Boolean
        objProduit = New Produit
        Try
            If pid <> 0 Then
                bReturn = objProduit.load(pid)
                If Not bReturn Then
                    setError("Produit.createAndLoad", getErreur())
                End If

            End If
        Catch ex As Exception
            setError("Produit.createAndLoad", ex.ToString)
        End Try
        Debug.Assert(objProduit.id = pid, "Produit " & pid & " non chargée")
        Return objProduit
    End Function 'createandload
    ''' <summary>
    ''' Constructeur pour Chargement par la clé
    ''' </summary>
    ''' <param name="pCode"> Code Produit</param>
    ''' <returns>Objet Produit ou null</returns>
    ''' <remarks></remarks>
    Public Shared Function createandloadbyKey(ByVal pCode As String) As Produit
        Dim objPRD As Produit
        Dim bReturn As Boolean
        Dim nId As Integer
        objPRD = New Produit
        Try
            If Not String.IsNullOrEmpty(pCode) Then
                nId = getIDPRDByKey(pCode)
                If nId <> -1 Then
                    bReturn = objPRD.load(nId)
                    If Not bReturn Then
                        setError("Produit.createAndLoad", getErreur())
                        objPRD = Nothing
                    End If
                Else
                    setError("Produit.createAndLoad", "No ID for " & pCode)
                    objPRD = Nothing
                End If
            End If
        Catch ex As Exception
            setError("Fournisseur.createAndLoad", ex.ToString)
            objPRD = Nothing
        End Try
        Return objPRD
    End Function 'Createanload

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
    Public Property nom() As String
        Get
            Return m_nom
        End Get
        Set(ByVal Value As String)
            If Value <> m_nom Then
                RaiseUpdated()
                m_nom = Value
            End If
        End Set
    End Property
    Public Property motcle() As String
        Get
            Return m_motcle
        End Get
        Set(ByVal Value As String)
            If Value <> m_motcle Then
                RaiseUpdated()
                m_motcle = Value
            End If
        End Set
    End Property
    Public Property millesime() As Integer
        Get
            Return m_millesime
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_millesime Then
                RaiseUpdated()
                m_millesime = Value
            End If
        End Set
    End Property
    Public Property idCouleur() As Integer
        Get
            Return m_idCouleur
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idCouleur Then
                RaiseUpdated()
                m_idCouleur = Value
            End If
        End Set
    End Property
    Public Property libCouleur() As String
        Get
            Dim objParam As Param
            Dim strReturn As String = String.Empty
            For Each objParam In Param.colCouleur
                If objParam.id = idCouleur Then
                    strReturn = objParam.valeur
                End If
            Next
            Return strReturn
        End Get
        Set(ByVal Value As String)
        End Set

    End Property
    Public Property idContenant() As Integer
        Get
            Return m_idContenant
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idContenant Then
                RaiseUpdated()
                m_idContenant = Value
            End If
        End Set
    End Property
    Public Property libContenant() As String
        Get
            Dim objParam As contenant
            Dim strReturn As String = String.Empty
            For Each objParam In contenant.colContenant
                If objParam.id = idContenant Then
                    strReturn = objParam.libelle
                End If
            Next
            Return strReturn
        End Get
        Set(ByVal Value As String)
            m_libContenant = Value
        End Set

    End Property
    Public Property idConditionnement() As Integer
        Get
            Return m_idConditionnement
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idConditionnement Then
                bUpdated = True
                m_idConditionnement = Value
            End If
        End Set

    End Property
    Public Property libConditionnement() As String
        Get
            Dim objParam As Param
            Dim strReturn As String = String.Empty
            For Each objParam In Param.colConditionnement
                If objParam.id = idConditionnement Then
                    strReturn = objParam.valeur
                End If
            Next
            Return strReturn

        End Get
        Set(ByVal Value As String)
        End Set
    End Property
    Public Property idRegion() As Integer
        Get
            Return m_idRegion
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idRegion Then
                RaiseUpdated()
                m_idRegion = Value
            End If
        End Set
    End Property
    Public Property libRegion() As String
        Get
            Dim objParam As Param
            Dim strReturn As String = String.Empty
            For Each objParam In Param.colRegion
                If objParam.id = idRegion Then
                    strReturn = objParam.valeur
                End If
            Next
            Return strReturn
        End Get
        Set(ByVal Value As String)
        End Set
    End Property
    Public Property bDisponible() As Boolean
        Get
            Return m_bDisponible
        End Get
        Set(ByVal Value As Boolean)
            If Value <> m_bDisponible Then
                RaiseUpdated()
                m_bDisponible = Value
            End If

        End Set
    End Property
    Public Property bStock() As Boolean
        Get
            Return m_bStock
        End Get
        Set(ByVal Value As Boolean)
            If Value <> m_bStock Then
                RaiseUpdated()
                m_bStock = Value
            End If
        End Set
    End Property
    Public Property idTVA() As Integer
        Get
            Return m_idTVA
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idTVA Then
                RaiseUpdated()
                m_idTVA = Value
            End If
        End Set
    End Property
    Public Property libTVA() As String
        Get
            Dim objParam As Param
            Dim strReturn As String = String.Empty
            For Each objParam In Param.colTVA
                If objParam.id = idCouleur Then
                    strReturn = objParam.valeur
                End If
            Next
            Return strReturn
        End Get
        Set(ByVal Value As String)
        End Set
    End Property
    Public Property codeStat() As String
        Get
            Return m_codeStat
        End Get
        Set(ByVal Value As String)
            If Value <> m_codeStat Then
                RaiseUpdated()
                m_codeStat = Value
            End If
        End Set
    End Property
    Public Property QteStock() As Decimal
        Get
            Return m_QteStock
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_QteStock Then
                RaiseUpdated()
                m_QteStock = Value
            End If
        End Set
    End Property
    Public Property DateDernInventaire() As Date
        Get
            Return m_DateDernInventaire
        End Get
        Set(ByVal Value As Date)
            If Value.ToShortDateString <> m_DateDernInventaire.ToShortDateString Then
                RaiseUpdated()
                m_DateDernInventaire = Value.ToShortDateString
            End If
        End Set
    End Property
    Public Property QteStockDernInventaire() As Double
        Get
            Return m_QteStockDernInventaire
        End Get
        Set(ByVal Value As Double)
            If Value <> m_QteStockDernInventaire Then
                RaiseUpdated()
                m_QteStockDernInventaire = Value
            End If
        End Set
    End Property
    Public Property idFournisseur() As Integer
        Get
            Return m_idFournisseur
        End Get
        Set(ByVal Value As Integer)
            If Value <> m_idFournisseur Then
                RaiseUpdated()
                m_idFournisseur = Value
            End If
        End Set
    End Property
    Public ReadOnly Property Tarif(ByVal strCode As String) As Decimal
        Get
            Select Case strCode
                Case "A"
                    Return TarifA
                Case "B"
                    Return TarifB
                Case "C"
                    Return TarifC
            End Select
        End Get
    End Property
    Public Property TarifA() As Decimal
        Get
            Return m_TarifA
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_TarifA Then
                RaiseUpdated()
                m_TarifA = Value
            End If
        End Set
    End Property
    Public Property TarifB() As Decimal
        Get
            Return m_TarifB
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_TarifB Then
                RaiseUpdated()
                m_TarifB = Value
            End If
        End Set
    End Property
    Public Property TarifC() As Decimal
        Get
            Return m_TarifC
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_TarifC Then
                RaiseUpdated()
                m_TarifC = Value
            End If
        End Set
    End Property
    Public Property nomFournisseur() As String
        Get
            Return m_nomFournisseur
        End Get
        Set(ByVal Value As String)

            m_nomFournisseur = Value
        End Set

    End Property
    Public ReadOnly Property qteCommande() As Decimal
        Get
            Return m_qteCommande
        End Get
    End Property 'qteCommande
    Friend Sub setQteCommande(ByVal Value As Decimal)
        m_qteCommande = Value
    End Sub 'SetQteCommande

    Public ReadOnly Property poids() As Decimal
        Get
            Dim objCont As contenant = Nothing
            Dim bok As Boolean
            bok = False
            For Each objCont In contenant.colContenant
                If objCont.id = m_idContenant Then
                    bok = True
                    Exit For
                End If
            Next
            If bok Then
                Return objCont.poids
            Else
                Debug.Assert(False, "Contenant non trouvé")
                Return 0
            End If

        End Get
    End Property
    ''' <summary>
    ''' Convertit une Quantité en nombre de colis en fonction du conditionnement
    ''' </summary>
    ''' <param name="nqte"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property qteColis(ByVal nqte As Decimal) As Integer
        Get
            Dim objCond As Param = Nothing
            Dim bok As Boolean
            Dim nColis As Integer = 0
            Try

                bok = False
                For Each objCond In Param.colConditionnement
                    If objCond.id = m_idConditionnement Then
                        bok = True
                        Exit For
                    End If
                Next
                If bok Then
                    nColis = nqte \ CDec(objCond.valeur) 'Rend la partie entière !!!
                    If nqte Mod CDec(objCond.valeur) <> 0 Then
                        'S'il y a un reste => ajout d'un colis
                        If nqte < 0 Then
                            nColis = nColis - 1
                        Else
                            nColis = nColis + 1
                        End If
                    End If

                    Return nColis

                Else
                    Debug.Assert(False, "Conditionnement non trouvé")
                    Return 0
                End If
            Catch ex As Exception
                Debug.Assert(False, "Produit.qteColis ERR", ex.Message)
                Return 0
            End Try

        End Get
    End Property
    Public Property DossierProduit() As String
        Get
            Return m_Dossier
        End Get
        Set(ByVal value As String)
            RaiseUpdated()
            m_Dossier = value
        End Set
    End Property
    Public ReadOnly Property colmvtStock() As ColEventSorted
        Get
            Return m_colMvtStock
        End Get
    End Property
    Public Sub releasecolmvtStock()
        m_colMvtStock = New ColEventSorted
        m_bcolMvtStockLoaded = False
        m_bcolMvtStockUpdated = False
    End Sub

    Private Sub m_colMvtStock_Updated() Handles m_colMvtStock.Updated
        RaiseUpdated()
    End Sub


#End Region
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Dim strReturn As String
        strReturn = m_code & "," & m_nom & "," & m_motcle & "," & m_millesime & "," & m_idCouleur & "," & libCouleur & "," & m_idContenant & "," & m_libContenant & "," & m_idConditionnement & "," & libConditionnement & "," & m_idRegion & "," & libRegion & "," & m_bDisponible & "," & m_idTVA & "," & libTVA & "," & m_codeStat & "," & m_QteStock & "," & m_DateDernInventaire & "," & m_QteStockDernInventaire & "," & m_idFournisseur & "," & m_nomFournisseur

        Return "PRD =(" & MyBase.toString() & strReturn & ")"
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objPrd As Produit
        Dim bReturn As Boolean
        Try
            objPrd = CType(obj, Produit)
            bReturn = MyBase.Equals(obj)

            bReturn = bReturn And (code.Equals(objPrd.code))
            bReturn = bReturn And (nom.Equals(objPrd.nom))
            bReturn = bReturn And (millesime.Equals(objPrd.millesime))
            bReturn = bReturn And (idCouleur.Equals(objPrd.idCouleur))
            bReturn = bReturn And (idContenant.Equals(objPrd.idContenant))
            bReturn = bReturn And (idConditionnement.Equals(objPrd.idConditionnement))
            bReturn = bReturn And (idRegion.Equals(objPrd.idRegion))
            bReturn = bReturn And (Me.idFournisseur.Equals(objPrd.idFournisseur))
            bReturn = bReturn And (Me.bDisponible.Equals(objPrd.bDisponible))
            bReturn = bReturn And (codeStat.Equals(objPrd.codeStat))
            bReturn = bReturn And (DateDernInventaire.Equals(objPrd.DateDernInventaire))
            bReturn = bReturn And (idTVA.Equals(objPrd.idTVA))
            bReturn = bReturn And (motcle.Equals(objPrd.motcle))
            bReturn = bReturn And (QteStock.Equals(objPrd.QteStock))
            bReturn = bReturn And (QteStockDernInventaire.Equals(objPrd.QteStockDernInventaire))
            bReturn = bReturn And (bStock.Equals(objPrd.bStock))
            bReturn = bReturn And (TarifA.Equals(objPrd.TarifA))
            bReturn = bReturn And (TarifB.Equals(objPrd.TarifB))
            bReturn = bReturn And (TarifC.Equals(objPrd.TarifC))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
#Region "Interface Persist"
    '=======================================================================
    'Fonction : Load
    'Description : Chargement de l'objet en base
    'Détails    : Appelle LoadPRD (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        init()
        shared_connect()
        If pid <> 0 Then
            m_id = pid
        End If
        bReturn = loadPRD()
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            m_bNew = False
            setUpdatedFalse()
            m_bDeleted = False
            m_colMvtStock.clear()
            m_bcolMvtStockLoaded = False ' la collection des mvt de stock n'est pas chargée
        End If
        Return bReturn
    End Function 'dbLoad
    '=======================================================================
    'Fonction : DBLoadLight
    'Description : Chargement de l'objet en base
    'Détails    : Appelle LoadPRD (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Public Function DBLoadLight(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        init()
        shared_connect()
        If pid <> 0 Then
            m_id = pid
        End If
        bReturn = loadPRDLight()
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            m_bNew = False
            setUpdatedFalse()
            m_bDeleted = False
            m_colMvtStock.clear()
            m_bcolMvtStockLoaded = False ' la collection des mvt de stock n'est pas chargée
        End If
        Return bReturn
    End Function 'dbLoadLight
    '=======================================================================
    'Fonction : delete
    'Description : suppression de l'objet en base
    'Détails    : Appelle deletePRD (de Persist) 
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = deletePRDMVTStock() 'Suppression des Mvts de Stocks
        'On Considère que la collection n'est plus chargée
        m_colMvtStock = New ColEventSorted
        m_bcolMvtStockLoaded = False

        If bReturn Then
            bReturn = deletePRDPReCommande() 'Suppression des Lignes de Precommande
            If bReturn Then
                bReturn = deletePRD()
            End If
            shared_disconnect()
        End If
        Return bReturn
    End Function 'delete
    '=======================================================================
    'Fonction : checkFordelete
    'description : Controle si l'élément est supprimable
    '                table commandesClients
    '=======================================================================
    Public Overrides Function checkForDelete() As Boolean
        Dim bReturn As Boolean
        Try
            shared_connect()
            bReturn = True
            If existeLgCommandeProduit() Then
                bReturn = False
            End If
            shared_disconnect()
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete
    '=======================================================================
    'Fonction : Insert
    'Description : Insert de l'objet en base
    'Détails    : Appelle InsertPRD (de Persist) 
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = insertPRD()
        shared_disconnect()
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : update
    'Description : Update de l'objet en base
    'Détails    : Appelle UpdatePRD (de Persist) 
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = updatePRD()
        shared_disconnect()
        Return bReturn
    End Function 'update

    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        Dim bWasDeleted As Boolean
        bWasDeleted = bDeleted
        shared_connect()
        bReturn = MyBase.Save()
        If (bReturn And Not bWasDeleted) Then
            If m_bcolMvtStockUpdated Then
                bReturn = savecolmvtStock()
            End If
        End If
        shared_disconnect()
        Return bReturn
    End Function

#End Region
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_code & " | " & m_nom & " | " & m_millesime & " | " & libContenant & " | " & libCouleur
        End Get
    End Property
    ''' <summary>
    ''' Chargement de toute la Liste des mouvents de Stocks
    ''' </summary>
    ''' <returns>TRUE/FALSE</returns>
    ''' <remarks></remarks>
    Public Function loadcolmvtStock() As Boolean
        Dim bReturn As Boolean

        bReturn = True
        If Not m_bcolMvtStockLoaded Then
            shared_connect()
            m_bcolMvtStockLoaded = True ' La collection des mouvements de stock est considérée comme chargée
            m_colMvtStock = mvtStock.getListe(m_id)
            m_bcolMvtStockLoaded = bReturn ' La Liste des mouvement de stock est chargée
            m_bcolMvtStockUpdated = False ' La Liste des mouvements de stocks n'est pas mise à jour
            shared_disconnect()
        End If
        Return bReturn
    End Function 'loadcolmvtStock
    ''' <summary>
    ''' Chargement de la Liste des mouvents de Stocks depuis le dernier mouvement d'inventaire
    ''' </summary>
    ''' <returns>TRUE/FALSE</returns>
    ''' <remarks></remarks>
    Public Function loadcolmvtStockDepuisLeDernierMouvementInventaire() As Boolean
        Dim bReturn As Boolean

        bReturn = True
        m_colMvtStock.clear()
        shared_connect()
        m_bcolMvtStockLoaded = True ' La collection des mouvements de stock est considérée comme chargée
        m_colMvtStock.AddRange(mvtStock.getListeDepuisDernMvtInventaire(m_id))
        m_bcolMvtStockLoaded = bReturn ' La Liste des mouvement de stock est chargée
        m_bcolMvtStockUpdated = False ' La Liste des mouvements de stocks n'est pas mise à jour
        shared_disconnect()
        Return bReturn
    End Function 'loadcolmvtStock
    Public Function loadcolmvtStockFactureColisage(pIdFactCol As Integer) As Boolean
        Dim bReturn As Boolean

        bReturn = True
        m_colMvtStock.clear()
        shared_connect()
        m_bcolMvtStockLoaded = True ' La collection des mouvements de stock est considérée comme chargée
        m_colMvtStock.AddRange(mvtStock.getListeColisage(Me.id, pIdFactCol))
        m_bcolMvtStockLoaded = bReturn ' La Liste des mouvement de stock est chargée
        m_bcolMvtStockUpdated = False ' La Liste des mouvements de stocks n'est pas mise à jour
        shared_disconnect()
        Return bReturn
    End Function 'loadcolmvtStock
    '=======================================================================
    'Fonction : savecolmvtStock
    'Description : Sauvegarde des lignes de mouvements de Stocks
    'Détails    : 
    'Retour : 
    '=======================================================================
    Public Function savecolmvtStock() As Boolean
        Dim bReturn As Boolean
        Dim objMvtStock As mvtStock

        Debug.Assert(Not m_colMvtStock Is Nothing, "m_col <> nothing")
        Debug.Assert(m_id <> 0, "Le Produit doit être saubegardé au préalable")
        'On Parcours la liste des mouvements de stocks 
        Persist.shared_connect()
        bReturn = True
        For Each objMvtStock In colmvtStock
            bReturn = bReturn And objMvtStock.save()
        Next
        Persist.shared_disconnect()
        Return bReturn
    End Function

    '===========================================================
    'Function : ajouteLigneMvtStock
    'Description : Ajjoute une ligne de mouvement de Stock au produit
    '               Calcul le nouveau stock
    '==========================================================
    Public Function ajouteLigneMvtStock(ByVal pobjmvt As mvtStock, Optional ByVal pbCalCulStock As Boolean = True) As mvtStock

        Debug.Assert(m_bcolMvtStockLoaded, "La Collection des mouvements de stocks  doit être chargée au préalable")
        Debug.Assert(pobjmvt.idProduit = m_id, "Les ID de produits doivent être identiques")
        Try
            m_colMvtStock.Add(pobjmvt, pobjmvt.key)
            If pbCalCulStock Then
                calculStock(pobjmvt)
            End If
            setcolMvtStockUpdated()
            Return pobjmvt
        Catch ex As Exception
            setError("AjoutMvtStock", ex.ToString)
            Return Nothing
        End Try
    End Function ' ajouteLigneMvtStock
    '===========================================================
    'Function : ajouteLigneMvtStock
    'Description : Créé une ligne de mouvement de stock et l'ajoute à la collection
    '               Calcul le nouveau stock
    '==========================================================
    Public Function ajouteLigneMvtStock(ByVal pdatemvt As Date, ByVal ptypemvt As vncTypeMvt, ByVal prefid As Integer, ByVal plib As String, ByVal pqte As Decimal, Optional ByVal pcom As String = "", Optional ByVal pbCalculStock As Boolean = True) As mvtStock
        Dim objMvtStock As mvtStock

        Debug.Assert(m_bcolMvtStockLoaded, "La Collection des mouvements de stocks  doit être chargée au préalable")
        Try
            objMvtStock = New mvtStock(pdatemvt, Me.id, ptypemvt, pqte, plib)
            objMvtStock.Commentaire = pcom
            objMvtStock.idReference = prefid

            objMvtStock = ajouteLigneMvtStock(objMvtStock, pbCalculStock)
            setcolMvtStockUpdated()
            Return objMvtStock
        Catch ex As Exception
            setError("AjoutMvtStock", ex.ToString)
            Return Nothing
        End Try
    End Function ' ajouteLigneMvtStock  
    '===========================================================
    'Function : genereMvtInventaire
    'Description : Créé une ligne d'inventaire avec la Qte en Stock
    '               Calcul le nouveau stock
    '==========================================================
    Public Function genereMvtInventaire(ByVal pDateMvt As Date, ByVal pQte As Decimal) As Boolean
        Debug.Assert(m_bcolMvtStockLoaded, "La Collection des mouvements de stocks  doit être chargée au préalable")
        Dim bReturn As Boolean
        Try
            ajouteLigneMvtStock(pDateMvt, vncEnums.vncTypeMvt.vncMvtInventaire, 0, "Mvt généré", pQte, "Mvt généré le " + DateAndTime.Now(), True)
            bReturn = True
        Catch ex As Exception
            setError("genereMvtInventaire", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function
    '===========================================================
    'Function : SupprimeLigneMvtStock
    'Description : Supprime une ligne de mvouvement de Stock
    '               Calcul le nouveau stock
    '==========================================================
    Public Function supprimeLigneMvtStock(ByVal pnligne As Integer, Optional ByVal pbCalculStock As Boolean = True) As mvtStock
        Dim objMvtStock As mvtStock = Nothing
        'Dim breturn As Boolean

        Debug.Assert(m_bcolMvtStockLoaded, "La Collection des mouvements de stocks  doit être chargée au préalable")
        Debug.Assert(m_colMvtStock.Count >= pnligne, "Le Numéro de ligne doit être <= au nombre de ligne")
        Try
            objMvtStock = m_colMvtStock(pnligne)
            objMvtStock.bDeleted = True
            'breturn = m_colMvtStock.Remove(pnligne) // ne pas supprimer de la collection , le marquer comme à supprimer uniquement
            If pbCalculStock Then
                recalculStock() ' il faut relancer le calcul complet car on a pu supprimer une ligne d'inventaire 
            End If
            setcolMvtStockUpdated()
            Return objMvtStock
        Catch ex As Exception
            setError("SupprimeLigneMvtStock", ex.ToString)
            Return Nothing
        End Try
    End Function ' SupprimeLigneMvtStock
    '===========================================================
    'Function : SupprimeLigneMvtStockCommande
    'Description : Supprime les lignes de mvouvement de Stock d'une commande
    '               Calcul le nouveau stock
    '==========================================================
    Public Function supprimeLigneMvtStock(ByVal pidCommande As Integer, ByVal pTypeMvt As vncTypeMvt, Optional ByVal pbCalculStock As Boolean = True) As Boolean
        Dim objMvtStock As mvtStock
        Dim breturn As Boolean
        Dim bFin As Boolean
        'Dim i As Integer

        Debug.Assert(m_bcolMvtStockLoaded, "La Collection des mouvements de stocks  doit être chargée au préalable")
        Debug.Assert(pidCommande <> 0, "L'id Commande doit être renseigné")
        Try
            'Boucle de Suprression des lignes de mvtde stocks
            For Each objMvtStock In colmvtStock
                If objMvtStock.idReference = pidCommande And objMvtStock.typeMvt = pTypeMvt Then
                    objMvtStock.bDeleted = True
                    bFin = False
                End If
            Next

            If pbCalculStock Then
                breturn = recalculStock() ' il faut relancer le calcul complet car on a pu supprimer une ligne d'inventaire 
            End If
            setcolMvtStockUpdated()
            Return breturn
        Catch ex As Exception
            setError("SupprimeLigneMvtStockCommande", ex.ToString)
            breturn = False
        End Try
        Debug.Assert(breturn, "SupprimeLigneMvtCommande" & getErreur())
        Return breturn
    End Function ' SupprimeLigneMvtStockCommande
    Public Sub setcolMvtStockUpdated()
        m_bcolMvtStockUpdated = True
    End Sub
    '===========================================================
    'Function : CalculStock
    'Description : Calcul du stock Produit
    '               Si la date du mouvement est >= à la date du dernier inventaire
    '                   Si c'est in mvt d'inventaire
    '                       MAJ des information de stocks
    '                   Sinon
    '                       MAJ du stock produit  
    '               Cette fonction suppose que les mvts de stocks arrivent triés par ordre croissant de date
    '==========================================================
    Public Function calculStock(ByVal pobjMvtStock As mvtStock) As Boolean
        Dim bReturn As Boolean
        Try
            If pobjMvtStock.datemvt >= DateDernInventaire Then
                If pobjMvtStock.typeMvt = vncEnums.vncTypeMvt.vncMvtInventaire Then
                    'Cas d'un mouvement d'inventaire
                    'Il faut recalculer le stock car on ne sait pas quel sont les mvt à prendre en compte
                    '                    QteStock = pobjMvtStock.qte
                    '                    QteStockDernInventaire = pobjMvtStock.qte
                    '                   DateDernInventaire = pobjMvtStock.datemvt
                    recalculStock()
                Else
                    'Autres type de mvt
                    QteStock = QteStock + pobjMvtStock.qte
                End If
                bReturn = True
            End If
        Catch ex As Exception
            setError("CalculStock", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function
    '===========================================================
    'Function : reCalculStock
    'Description : reCalcul du stock Produit
    '                   Recherche du dernier inventaire
    '                   Parcours des mvts de stocks depuis
    '                   Mise à jour des informations de stocks
    '==========================================================
    Public Function recalculStock() As Boolean
        Dim bReturn As Boolean
        Dim nIndexDernierMvtInventaire As Integer
        Dim objMvtDernierInventaire As mvtStock = Nothing
        Dim nStockDepart As Decimal
        Dim nStock As Decimal
        Dim i As Integer
        Dim objmvtStock As mvtStock = Nothing

        Try
            'Recherche du dernier inventaire
            nIndexDernierMvtInventaire = rechercheDernierInventaire()
            If nIndexDernierMvtInventaire <> -1 Then
                Debug.Assert(nIndexDernierMvtInventaire <= m_colMvtStock.Count)
                objMvtDernierInventaire = m_colMvtStock(nIndexDernierMvtInventaire)
                nStockDepart = objMvtDernierInventaire.qte
            Else
                nStockDepart = 0
            End If
            'Parcours des mvt de stock depuis
            nStock = nStockDepart
            'S'il n'y a pas de mvts d'inventaire on parcours toute la table
            If nIndexDernierMvtInventaire = -1 Then
                nIndexDernierMvtInventaire = m_colMvtStock.Count
            End If
            'Si le mvt d'inventaire est le mvt le plus récent , on ne parcours pas la liste
            If nIndexDernierMvtInventaire > 0 Then
                For i = nIndexDernierMvtInventaire - 1 To 0 Step -1
                    objmvtStock = m_colMvtStock(i)
                    If (Not objmvtStock.bDeleted) Then
                        nStock = nStock + objmvtStock.qte
                    End If
                Next
            End If
            'Mise a jour du stock produit
            QteStock = nStock
            QteStockDernInventaire = nStockDepart
            If objMvtDernierInventaire Is Nothing Then
                DateDernInventaire = DATE_DEFAUT
            Else
                DateDernInventaire = objMvtDernierInventaire.datemvt
            End If
            bReturn = True
        Catch ex As Exception
            setError("RecalculStock", ex.ToString)
            bReturn = False
        End Try

        Return bReturn
    End Function 'recalculStock

    '===========================================================
    'Function : CalculStockAu
    'Description : Calcule le Stock au début du jour donnée
    '                   Recherche du dernier inventaire
    '                   Parcours des mvts de stocks depuis
    '                   Mise à jour des informations de stocks
    '==========================================================
    Public Function CalculeStockAu(ByVal pDate As Date) As Decimal
        Dim bReturn As Boolean
        Dim nIndexDernierMvtInventaire As Integer
        Dim objMvtDernierInventaire As mvtStock = Nothing
        Dim nStockDepart As Decimal
        Dim nStock As Decimal
        Dim i As Integer
        Dim objmvtStock As mvtStock = Nothing

        Try
            'Recherche du dernier inventaire
            loadcolmvtStockDepuisLeDernierMouvementInventaire()
            nIndexDernierMvtInventaire = rechercheDernierInventaire()
            If nIndexDernierMvtInventaire <> -1 Then
                Debug.Assert(nIndexDernierMvtInventaire <= m_colMvtStock.Count)
                objMvtDernierInventaire = m_colMvtStock(nIndexDernierMvtInventaire)
                nStockDepart = objMvtDernierInventaire.qte
            Else
                nStockDepart = 0
            End If
            'Parcours des mvt de stock depuis
            nStock = nStockDepart
            'S'il n'y a pas de mvts d'inventaire on parcours toute la table
            If nIndexDernierMvtInventaire = -1 Then
                nIndexDernierMvtInventaire = m_colMvtStock.Count
            End If
            'Si le mvt d'inventaire est le mvt le plus récent , on ne parcours pas la liste
            If nIndexDernierMvtInventaire > 0 Then
                For i = nIndexDernierMvtInventaire - 1 To 0 Step -1

                    objmvtStock = m_colMvtStock(i)
                    If Not objmvtStock.bDeleted Then
                        If objmvtStock.datemvt < pDate Then
                            nStock = nStock + objmvtStock.qte
                        End If
                    End If
                Next
            End If
            bReturn = True
        Catch ex As Exception
            setError("CalculStockAu", ex.ToString)
            bReturn = False
        End Try

        Return nStock
    End Function 'CalculeStockAu
    '============================================================
    ' Function : RechercheDernierInventaire
    ' Description : Recherche du mouvement inventaire le plus Récent
    ' Retour : Renvoie le numéro d'index dans la collection correspondant ou -1
    '===========================================================
    Private Function rechercheDernierInventaire() As Integer
        Debug.Assert(Not m_colMvtStock Is Nothing)
        Debug.Assert(m_bcolMvtStockLoaded, "La Collection des mouvements de stocks doit être chargée")

        Dim objMvt As mvtStock
        Dim objDernierInventaire As mvtStock
        Dim nReturn As Integer
        Dim i As Integer
        Try
            nReturn = -1
            For i = 0 To m_colMvtStock.Count - 1
                objMvt = CType(m_colMvtStock(i), mvtStock)
                If Not (objMvt.bDeleted) Then

                    If objMvt.typeMvt = vncTypeMvt.vncMvtInventaire Then
                        objDernierInventaire = objMvt
                        nReturn = i
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            nReturn = -1
        End Try

        Debug.Assert(nReturn <= m_colMvtStock.Count)
        Return nReturn
    End Function ' rechercheDernierInventaire
    '============================================================
    ' Function : PurgeMvtStock
    ' Description : Purge des mouvements de stock d'un produit et recalcul du stock
    ' Retour : Vrai ou faux
    '===========================================================
    Public Function purgeMvtStock(ByVal pDatePurge As Date) As Boolean
        Dim bReturn As Boolean
        Persist.shared_connect()
        Try
            bReturn = True
            bReturn = purgePRDMVTStock(pDatePurge)
            If bReturn Then
                m_colMvtStock = New ColEventSorted
                m_bcolMvtStockLoaded = False
                bReturn = loadcolmvtStock()
                If bReturn Then
                    bReturn = recalculStock()
                End If
            End If

        Catch ex As Exception
            bReturn = False
            setError("Produit.purge", ex.ToString())
        End Try

        Persist.shared_disconnect()
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'PurgeMvtStock


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Function Fill(objRS As OleDb.OleDbDataReader) As Boolean
        Dim bReturn As Boolean = True
        For n As Integer = 0 To objRS.FieldCount - 1
            If Not objRS.IsDBNull(n) Then
                bReturn = bReturn And Fill(objRS.GetName(n), objRS.GetValue(n))
            End If
        Next
        m_bResume = True
        Return bReturn
    End Function
    Public Function Fill(pColName As String, pColValue As Object) As Boolean
        Dim bReturn As Boolean = True
        Try

            Select Case pColName.ToUpper.Trim
                Case "PRD_ID"
                    Me.setid(Convert.ToInt32(pColValue))
                Case "PRD_CODE"
                    Me.code = pColValue.ToString()
                Case "PRD_DISPO"
                    Me.bDisponible = Convert.ToBoolean(pColValue)
                Case "PRD_STOCK"
                    Me.bStock = Convert.ToBoolean(pColValue)
                Case "PRD_CODE_STAT"
                    Me.codeStat = Convert.ToString(pColValue)
                Case "PRD_DATE_DERN_INVENT"
                    Me.DateDernInventaire = Convert.ToString(pColValue)
                Case "PRD_COND_ID"
                    Me.idConditionnement = Convert.ToInt32(pColValue)
                Case "PRD_CONT_ID"
                    Me.idContenant = Convert.ToInt32(pColValue)
                Case "PRD_COUL_ID"
                    Me.idCouleur = Convert.ToInt32(pColValue)
                Case "PRD_FRN_ID"
                    Me.idFournisseur = Convert.ToInt32(pColValue)
                Case "PRD_RGN_ID"
                    Me.idRegion = Convert.ToInt32(pColValue)
                Case "PRD_TVA_ID"
                    Me.idTVA = Convert.ToInt32(pColValue)
                Case "PRD_LIBELLE"
                    Me.nom = Convert.ToString(pColValue)
                Case "PRD_MIL"
                    Me.millesime = Convert.ToString(pColValue)
                Case "PRD_MOT_CLE"
                    Me.motcle = Convert.ToString(pColValue)
                Case "PRD_QTE_STK"
                    Me.QteStock = Convert.ToString(pColValue)
                Case "PRD_QTE_STOCK_DERN_INVENT"
                    Me.QteStockDernInventaire = Convert.ToString(pColValue)
                Case "PRD_TARIFA"
                    Me.TarifA = Convert.ToString(pColValue)
                Case "PRD_TARIFB"
                    Me.TarifB = Convert.ToString(pColValue)
                Case "PRD_TARIFC"
                    Me.TarifC = Convert.ToString(pColValue)
                Case "PRD_DOSSIER"
                    Me.DossierProduit = Convert.ToString(pColValue)
                Case "PRD_QTE_COMMANDE"
                    Me.setQteCommande(Convert.ToInt32(pColValue))
            End Select
        Catch ex As Exception
            setError("Produit.Fill(" & pColName & "," & pColValue.ToString() & ") ERR :" & ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function

    Public Function GenereDataSetRecapColisage(ByVal pdDeb As Date, ByVal pdFin As Date, ByVal pCout As Decimal, ByRef pDs As dsVinicom) As Boolean

        'Debug.Assert(Not String.IsNullOrEmpty(pCodeFourn), "Le Code Fournisseur doit être spécifié")
        Dim bReturn As Boolean
        Dim nSI As Decimal
        Dim nSF As Decimal
        Dim nEntree As Decimal
        Dim nSortie As Decimal
        Dim periode As String
        Dim nJourDansleMois As Integer = pdFin.Day

        Dim dDatePrec As Date = DateTime.MinValue
        periode = pdDeb.ToString("MMMM yyyy")
        nSI = 0
        nSF = 0
        nEntree = 0
        nSortie = 0
        bReturn = False
        Try
            Dim oPRD As Produit = Me
            Dim oStockAuQ As Decimal = oPRD.CalculeStockAu(pdDeb)  'Calcul du Stock la veille du Début
            Dim oStockAu As Integer = oPRD.qteColis(oStockAuQ) 'Conversion en Colis

            Dim nmvtDu(31) As Integer
            Dim nStockAu(31) As Integer
            For njour As Integer = 1 To 31
                nmvtDu(njour) = 0
                nStockAu(njour) = 0
            Next
            Dim lstmvtStock As List(Of mvtStock)
            lstmvtStock = mvtStock.regroupMvtStockmemecommande(oPRD.colmvtStock)
            'Boucle pour calculer les mvt par jour
            For nIndex As Integer = lstmvtStock.Count - 1 To 0 Step -1
                Dim oMvtStk As mvtStock = lstmvtStock(nIndex)
                If oMvtStk.datemvt >= pdDeb And oMvtStk.datemvt <= pdFin Then
                    'C'est un Mvt de la bonne date
                    If oMvtStk.typeMvt <> vncTypeMvt.vncMvtInventaire Then
                        Dim qteColis As Integer = oPRD.qteColis(oMvtStk.qte)
                        nmvtDu(oMvtStk.datemvt.Day) = nmvtDu(oMvtStk.datemvt.Day) + qteColis
                    End If
                End If
            Next
            'Boucle pour calculer le stock à la fin de chaque journée
            For njour As Integer = 1 To nJourDansleMois
                nStockAu(njour) = oStockAu + nmvtDu(njour)
                oStockAu = nStockAu(njour)
            Next

            Dim oFournisseur As Fournisseur
            If oPRD.DossierProduit = Dossier.VINICOM Then
                oFournisseur = Fournisseur.createandload(oPRD.idFournisseur)
            End If
            If oPRD.DossierProduit = Dossier.HOBIVIN Then
                oFournisseur = Fournisseur.getIntermediairePourUnDossier(oPRD.DossierProduit)
            End If
            pDs.RECAPCOLISAGEJOURN.AddRECAPCOLISAGEJOURNRow(RC_PRD_CODE:=oPRD.code, RC_PRD_LIBELLE:=oPRD.nom,
                                                            RC_FRN_CODE:=oFournisseur.code,
                                                            RC_FRN_NOM:=oFournisseur.nom,
                                                            RC_FRN_RS:=oFournisseur.rs,
                                                            RC_COUT_U:=pCout,
                                                            RC_S01:=nStockAu(1),
                                                            RC_S02:=nStockAu(2),
                                                            RC_S03:=nStockAu(3),
                                                            RC_S04:=nStockAu(4),
                                                            RC_S05:=nStockAu(5),
                                                            RC_S06:=nStockAu(6),
                                                            RC_S07:=nStockAu(7),
                                                            RC_S08:=nStockAu(8),
                                                            RC_S09:=nStockAu(9),
                                                            RC_S10:=nStockAu(10),
                                                            RC_S11:=nStockAu(11),
                                                            RC_S12:=nStockAu(12),
                                                            RC_S13:=nStockAu(13),
                                                            RC_S14:=nStockAu(14),
                                                            RC_S15:=nStockAu(15),
                                                            RC_S16:=nStockAu(16),
                                                            RC_S17:=nStockAu(17),
                                                            RC_S18:=nStockAu(18),
                                                            RC_S19:=nStockAu(19),
                                                            RC_S20:=nStockAu(20),
                                                            RC_S21:=nStockAu(21),
                                                            RC_S22:=nStockAu(22),
                                                            RC_S23:=nStockAu(23),
                                                            RC_S24:=nStockAu(24),
                                                            RC_S25:=nStockAu(25),
                                                            RC_S26:=nStockAu(26),
                                                            RC_S27:=nStockAu(27),
                                                            RC_S28:=nStockAu(28),
                                                            RC_S29:=nStockAu(29),
                                                            RC_S30:=nStockAu(30),
                                                            RC_S31:=nStockAu(31),
                                                            periode:=periode,
                                                            RC_IDPRODUIT:=oPRD.id
                                                            )




            bReturn = True
        Catch ex As Exception
            setError("Produit.GenereDataSetRecapcolisage", ex.Message)
            bReturn = False
        End Try


        Return bReturn

    End Function
    Public Function GenereDataSetRecapColisage(ByVal pIdFactCol As Integer, ByVal pCout As Decimal, ByRef pDs As dsVinicom) As Boolean

        Dim bReturn As Boolean
        Dim nSI As Decimal
        Dim nSF As Decimal
        Dim nEntree As Decimal
        Dim nSortie As Decimal
        Dim periode As String
        Dim oFactCol As FactColisageJ
        oFactCol = FactColisageJ.createandload(pIdFactCol)
        Dim dDeb As Date = CDate(oFactCol.periode)
        Dim dFin As Date = dDeb.AddMonths(1).AddDays(-1)
        Dim nJourDansleMois As Integer = dFin.Day

        Dim dDatePrec As Date = DateTime.MinValue
        periode = oFactCol.periode
        nSI = 0
        nSF = 0
        nEntree = 0
        nSortie = 0
        bReturn = False
        Try
            'La Liste des Mouvements de Stocks est chargée depuis une facture de colisage
            'Le Premier Element DOIT ETRE un Mouvement de type inventaire
            'NB La liste est chargée en send inverse des dates , c'est donc le dernier element de la liste

            Dim oStockAuQ As Decimal = 0
            If colmvtStock.Count > 0 Then
                Dim omvtIventaire As mvtStock
                omvtIventaire = colmvtStock.Item(colmvtStock.Count - 1)
                If omvtIventaire.typeMvt = vncTypeMvt.vncMvtInventaire Then
                    oStockAuQ = omvtIventaire.qte
                End If
            End If
            Dim oStockAu As Integer = qteColis(oStockAuQ) 'Conversion en Colis

            Dim nmvtDu(31) As Integer
            Dim nStockAu(31) As Integer
            For njour As Integer = 1 To 31
                nmvtDu(njour) = 0
                nStockAu(njour) = 0
            Next
            'Boucle pour calculer les mvt par jour
            For nIndex As Integer = colmvtStock.Count - 1 To 0 Step -1
                Dim oMvtStk As mvtStock = colmvtStock(nIndex)
                If oMvtStk.idFactColisage = pIdFactCol Then
                    'C'est un Mvt de la bonne Facture
                    If oMvtStk.typeMvt <> vncTypeMvt.vncMvtInventaire Then
                        Dim nbColis As Integer = qteColis(oMvtStk.qte)
                        nmvtDu(oMvtStk.datemvt.Day) = nmvtDu(oMvtStk.datemvt.Day) + nbColis
                    End If
                End If
            Next
            'Boucle pour calculer le stock à la fin de chaque journée
            For njour As Integer = 1 To nJourDansleMois
                nStockAu(njour) = oStockAu + nmvtDu(njour)
                oStockAu = nStockAu(njour)
            Next

            Dim oFournisseur As New Fournisseur()
            If DossierProduit = Dossier.VINICOM Then
                oFournisseur = Fournisseur.createandload(idFournisseur)
            End If
            If DossierProduit = Dossier.HOBIVIN Then
                oFournisseur = Fournisseur.getIntermediairePourUnDossier(DossierProduit)
            End If
            pDs.RECAPCOLISAGEJOURN.AddRECAPCOLISAGEJOURNRow(RC_PRD_CODE:=code, RC_PRD_LIBELLE:=nom,
                                                            RC_FRN_CODE:=oFournisseur.code,
                                                            RC_FRN_NOM:=oFournisseur.nom,
                                                            RC_FRN_RS:=oFournisseur.rs,
                                                            RC_COUT_U:=pCout,
                                                            RC_S01:=nStockAu(1),
                                                            RC_S02:=nStockAu(2),
                                                            RC_S03:=nStockAu(3),
                                                            RC_S04:=nStockAu(4),
                                                            RC_S05:=nStockAu(5),
                                                            RC_S06:=nStockAu(6),
                                                            RC_S07:=nStockAu(7),
                                                            RC_S08:=nStockAu(8),
                                                            RC_S09:=nStockAu(9),
                                                            RC_S10:=nStockAu(10),
                                                            RC_S11:=nStockAu(11),
                                                            RC_S12:=nStockAu(12),
                                                            RC_S13:=nStockAu(13),
                                                            RC_S14:=nStockAu(14),
                                                            RC_S15:=nStockAu(15),
                                                            RC_S16:=nStockAu(16),
                                                            RC_S17:=nStockAu(17),
                                                            RC_S18:=nStockAu(18),
                                                            RC_S19:=nStockAu(19),
                                                            RC_S20:=nStockAu(20),
                                                            RC_S21:=nStockAu(21),
                                                            RC_S22:=nStockAu(22),
                                                            RC_S23:=nStockAu(23),
                                                            RC_S24:=nStockAu(24),
                                                            RC_S25:=nStockAu(25),
                                                            RC_S26:=nStockAu(26),
                                                            RC_S27:=nStockAu(27),
                                                            RC_S28:=nStockAu(28),
                                                            RC_S29:=nStockAu(29),
                                                            RC_S30:=nStockAu(30),
                                                            RC_S31:=nStockAu(31),
                                                            PERIODE:=periode,
                                                            RC_IDPRODUIT:=id
                                                            )




            bReturn = True
        Catch ex As Exception
            setError("Produit.GenereDataSetRecapcolisage", ex.Message)
            bReturn = False
        End Try


        Return bReturn

    End Function



End Class
