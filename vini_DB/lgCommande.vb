'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : lgCommande
' Description : Ligne de Commande
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
'17/10/05 : Ajout du poids
'
'===================================================================================================================================
Imports System.Collections.Generic
Imports System.IO
Imports System.Xml.Serialization
Public Class LgCommande
    Inherits Persist
    Implements ICloneable

    Private m_num As Integer                'Numéro d'ordre de ligne
    Private m_oProduit As Produit           'Produit Commandé
    Private m_idCmd As Integer                 'Id de la commande
    Private m_idSCmd As Integer                 'Id de la Sous-commande
    Private m_qteCom As Decimal         'Quantite commandée
    Private m_qteLiv As Decimal         'Quantite commandée
    Private m_qteFact As Decimal         'Quantite commandée
    Private m_prixU As Decimal               'Prix Unitaire
    Private m_prixHT As Decimal               'Prix Total HT
    Private m_prixTTC As Decimal               'Prix Total TTC
    Private m_bGratuit As Boolean           'Ligne Gratuite
    Private m_bProduitLoaded As Boolean
    Private m_bLigneEclatee As Boolean
    Private m_idBA As Integer               'Id Du Bon Appro
    Private m_poids As Decimal         'Poids de la commande
    Private m_qteColis As Decimal         'Quantite de colis
    Private m_TxComm As Decimal         'Taux de commisson,
    Private m_MtComm As Decimal         'Montant de la Commsssion
    Private m_objSCMD As SousCommande    'sous commande (utilisé dans les ligne de factures)
    Private m_bStockDispo As Boolean    'Le Stock est dispoible (Par défaut oui)
    Private m_idTiers As Integer 'Identifiant du Tiers

#Region "Accesseurs"
    Protected Sub New()

    End Sub
    Public Sub New(ByVal pidCommande As Integer, Optional ByVal pidScmd As Integer = 0, Optional ByVal pidBA As Integer = 0)
        m_typedonnee = vncEnums.vncTypeDonnee.LGCOMMANDE
        m_idSCmd = 0

        If pidCommande <> 0 Then
            m_idCmd = pidCommande
        End If
        If pidScmd <> 0 Then
            m_idSCmd = pidScmd
        End If
        If pidBA <> 0 Then
            m_idBA = pidBA
        End If
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
        m_bLigneEclatee = False
        m_TxComm = 0.0
        m_MtComm = 0.0
        m_bStockDispo = True
        m_idTiers = 0
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
    Public Property idCmd() As Integer
        Get
            Return m_idCmd
        End Get
        Set(ByVal Value As Integer)
            If m_idCmd <> Value Then
                m_idCmd = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property idSCmd() As Integer
        Get
            Return m_idSCmd
        End Get
        Set(ByVal Value As Integer)
            If m_idSCmd <> Value Then
                m_idSCmd = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property idBA() As Integer
        Get
            Return m_idBA
        End Get
        Set(ByVal Value As Integer)
            If m_idBA <> Value Then
                m_idBA = Value
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
    Public Property bLigneEclatee() As Boolean
        Get
            Return (m_bLigneEclatee)
        End Get
        Set(ByVal Value As Boolean)
            If m_bLigneEclatee <> Value Then
                m_bLigneEclatee = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property poids() As Decimal
        Get
            Return m_poids
        End Get
        Set(ByVal Value As Decimal)
            If m_poids <> Value Then
                m_poids = Value
                RaiseUpdated()
            End If

        End Set
    End Property

    Public Property qteColis() As Decimal
        Get
            Return m_qteColis
        End Get
        Set(ByVal Value As Decimal)
            If m_qteColis <> Value Then
                m_qteColis = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    ''' <summary>
    ''' Taux de commission
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TxComm() As Decimal
        Get
            Return m_TxComm
        End Get
        Set(ByVal Value As Decimal)
            If m_TxComm <> Value Then
                m_TxComm = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Montant de commission
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MtComm() As Decimal
        Get
            Return m_MtComm
        End Get
        Set(ByVal Value As Decimal)
            If m_MtComm <> Value Then
                m_MtComm = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    'Propriétés utilisées lors de la gestion des factures de commission et de transport
    Public ReadOnly Property LGFACT_SCMD_CodeCommandeClient() As String
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.codeCommandeClient
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property LGFACT__SCMD_RSClient() As String
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.oClient.rs
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property LGFACT__SCMD_CodeFactFournisseur() As String
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.refFactFournisseur
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property LGFACT__SCMD_DateFactFournisseur() As Date
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.dateFactFournisseur
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property LGFACT__SCMD_MontantFactFournisseur() As Decimal
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.totalHTFacture
            Else
                Return 0.0
            End If
        End Get
    End Property

    Public ReadOnly Property LGFACT__SCMD_BaseCommission() As Decimal
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.baseCommission
            Else
                Return 0.0
            End If
        End Get
    End Property
    Public ReadOnly Property LGFACT__SCMD_MontantCommission() As Decimal
        Get
            If m_objSCMD Is Nothing Then
                m_objSCMD = SousCommande.createandload(m_idSCmd)
            End If
            If Not m_objSCMD Is Nothing Then
                Return m_objSCMD.MontantCommission
            Else
                Return 0.0
            End If
        End Get
    End Property
    ''' <summary>
    ''' Le Stock est-il disponible ?
    ''' Donnée non sauvegardée (recalculée à la demande)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property bStockDispo() As Boolean
        Get
            Return (m_bStockDispo)
        End Get
        Set(ByVal Value As Boolean)
            m_bStockDispo = Value
        End Set
    End Property
    ''' <summary>
    ''' Id du tiers associé à la commande
    ''' MAJ par le Load de la commande
    ''' non sauvegardé
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property idTiers() As Integer
        Get
            Return m_idTiers
        End Get
        Set(ByVal Value As Integer)
            m_idTiers = Value
        End Set
    End Property


    Public Sub calcPoidsColis()
        Debug.Assert(m_oProduit.id <> 0, "Produit inexistant")
        If qteLiv <> 0 Then
            calcPoidsColis(qteLiv)
        Else
            calcPoidsColis(qteCommande)
        End If
    End Sub
    Public Sub calcPoidsColis(pQte As Decimal)
        Debug.Assert(m_oProduit.id <> 0, "Produit inexistant")
        poids = m_oProduit.poids * pQte
        qteColis = m_oProduit.qteColis(pQte)
    End Sub

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
        Return "LGCM =(" & ")"
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
            bReturn = bReturn And (idCmd = (objLgCommande.idCmd))
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
            bReturn = bReturn And (idSCmd = objLgCommande.idSCmd)
            bReturn = bReturn And (poids = objLgCommande.poids)
            bReturn = bReturn And (qteColis = objLgCommande.qteColis)
            bReturn = bReturn And (TxComm = objLgCommande.TxComm)
            bReturn = bReturn And (MtComm = objLgCommande.MtComm)


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
            For Each oParam In Param.colTVA
                If oParam.id = oProduit.idTVA Then
                    nTaux = CDbl(oParam.valeur)
                    Exit For
                End If
            Next
            Debug.Assert(nTaux <> 0)

            If qteFact <> 0 Then
                qte = qteFact
            Else
                If qteLiv <> 0 Then
                    qte = qteLiv
                Else
                    qte = qteCommande
                End If
            End If
            m_prixHT = m_prixU * qte
            m_prixTTC = Math.Round(m_prixHT * (1 + (nTaux / 100)), 2)
        End If
        Return True
    End Function 'calculPrixTotal

    'Methode : EstLivree
    'Rend vrai si la ligne est Livrée

    Public Function estLivree() As Boolean
        Dim bReturn As Boolean
        'Une ligne est considérée comme livrée si 
        '    La Qte Livrée est <> 0
        '    La Qte Livrée est =  0 et Qte Commandée = 0
        '    La Qte Livrée est =  0 et Qte Facturée <> 0

        bReturn = False
        If m_qteLiv <> 0 Then
            bReturn = True
        Else
            If m_qteCom = 0 Then
                bReturn = True
            Else
                If m_qteFact <> 0 Then
                    bReturn = True
                End If
            End If
        End If
        Return bReturn
    End Function 'estLivree

    ''' <summary>
    ''' Calcul de la commission en fonction de la Qte
    ''' </summary>
    ''' <param name="pQteCalcul">Calculer la commsion en fonction de la qte Commandée ou Livrée</param>
    ''' <remarks></remarks>
    Public Sub CalculCommission(pstrOrigine As String, ByVal pQteCalcul As vncEnums.CalculCommQte)
        Dim nBaseCalul As Decimal

        'Récupération du Tx de comm s'il n'a pas déjà été calculé
        If m_TxComm = 0 Then
            m_TxComm = getTxComm(pstrOrigine)
        End If

        If pQteCalcul = CalculCommQte.CALCUL_COMMISSION_QTE_LIVREE Then
            nBaseCalul = qteLiv * prixU
        Else
            nBaseCalul = qteCommande * prixU
        End If

        m_MtComm = Math.Round(nBaseCalul * (m_TxComm / 100), 2)
    End Sub

    Private Function getTxComm(pstrOrigine As String) As Decimal
        Dim oTauxCommm As TauxComm
        Dim nReturn As Decimal
        Dim nIdFRn As Integer
        Dim strTypeClt As String

        Try

            nIdFRn = 0
            strTypeClt = String.Empty
            nReturn = 0.0
            'Get ID FRN
            '==========
            If m_oProduit IsNot Nothing Then
                nIdFRn = m_oProduit.idFournisseur
            End If
            'Get TypeClient
            '==============
            strTypeClt = getTypeClient(pstrOrigine)            'Récupératino du taux de commision
            '================================================
            nReturn = 0
            If Not String.IsNullOrEmpty(strTypeClt) Then
                oTauxCommm = TauxComm.getTauxComm(nIdFRn, strTypeClt)
                If oTauxCommm IsNot Nothing Then
                    nReturn = oTauxCommm.TauxComm
                End If
            End If
        Catch ex As Exception
            nReturn = 0
        End Try

        Return nReturn
    End Function
    ''' <summary>
    ''' Récupération du taux de commission en fonction de l'origine de la commande et du dossier du Produit
    ''' </summary>
    ''' <param name="pstrOrigine"></param>
    ''' <returns></returns>
    Private Function getTypeClient(pstrOrigine As String) As String
        Dim objCMD As CommandeClient
        Dim objCLT As Client
        Dim objParam As Param
        Dim strTypeClt As String = ""

        Select Case pstrOrigine
            Case Dossier.VINICOM
                'COMMANDE VINICOM
                Select Case m_oProduit.Dossier
                    Case Dossier.VINICOM
                        'un Produit Vinicom sur une Commande Vinicom
                        'On prend le type du client Final
                        If idTiers = 0 Then
                            'Load CMD
                            objCMD = CommandeClient.createandload(m_idCmd)
                            If objCMD IsNot Nothing Then
                                'Load CMD.CLT
                                objCLT = Client.createandload(objCMD.oTiers.id)
                            End If
                        Else
                            objCLT = Client.createandload(idTiers)
                        End If

                    Case Else
                        'pas de calcul de commission
                        objCLT = Nothing
                End Select
            Case Else
                'Commande HOBIVIN
                Select Case m_oProduit.Dossier
                    Case Dossier.VINICOM
                        'un Produit Vinicom sur une Commande HOBIVIN
                        'On prend le type du client Intermédiare
                        objCLT = Client.getIntermediairePourUneOrigine(pstrOrigine)
                    Case Else
                        'un produit HOBIVIN sur une commande HOBIVIN
                        objCLT = Nothing
                End Select
        End Select
        If objCLT IsNot Nothing Then
            ' Récupération du type du client
            objParam = New Param()
            objParam.load(objCLT.idTypeClient)
            strTypeClt = objParam.code
        Else
            'pas de calcul de commission
            strTypeClt = ""
        End If

        Return strTypeClt
    End Function

    ''' <summary>
    ''' Controle du stock disponible pour chaque ligne de commande
    ''' met à jour l'indicateur bStockDispo sur la ligne de commande
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function controleStockDispo() As Boolean
        Debug.Assert(Not oProduit Is Nothing, "Produit Non renseigné")
        Dim bReturn As Boolean
        Try
            Dim oPrd As New Produit()
            oPrd.load(oProduit.id)
            Dim nStockReel As Decimal = oPrd.QteStock
            Dim nQteComm As Decimal = oPrd.qteCommande - qteCommande
            Dim nStockTheo As Decimal = nStockReel - nQteComm

            Me.bStockDispo = nStockTheo >= qteCommande
            bReturn = True
        Catch ex As Exception
            setError("LgCommande.ControleStockdispo ERR : " & ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'controleStockDispo

    Function Clone() As Object Implements ICloneable.Clone
        Dim oSer As New XmlSerializer(GetType(LgCommande))
        Dim ostrW As StringWriter

        ostrW = New StringWriter()

        Me.oProduit.load(oProduit.id)
        oSer.Serialize(ostrW, Me)
        ostrW.Close()

        Dim strR As New StringReader(ostrW.ToString())

        Dim oLg As LgCommande
        oLg = TryCast(oSer.Deserialize(strR), LgCommande)
        oLg.majBooleenAlaFinDuNew()
        oLg.setid(0)
        oLg.oProduit.load(Me.oProduit.id) ' Rechargelment du produit
        Return oLg
    End Function


End Class
