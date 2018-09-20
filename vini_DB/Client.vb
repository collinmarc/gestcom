Option Explicit On 
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Client
' Description : Contient les informations du Client
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'   getListe    : Rend la Liste des Clients
'Protected
'Private
'Membres d'instances
'==================
'Public
'   idTypeClient : Identifiant du type de client
'   libTypeClient : (ReadOnly) Libellé du type de client
'   load(id) : Chargement de l'objet en base
'               Rend True si le chargement est OK
'   insert(id) : insertio de l'objet en base
'               Rend True si l'insertion  est OK
'   update(id) : MAJ de l'objet en base
'               Rend True si la MAJ est OK
'   delete(id) : Suppression de l'objet en base
'               Rend True si le DELETE est OK
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================
Imports System.Collections.Generic
Public Class Client
    Inherits Tiers
    Private m_IdPrestaShop As Long
    Private m_idTypeClient As Integer
    Private m_libTypeClient As String
    Private m_oPreCommande As preCommande
    Private m_CodeTarif As String
    Private m_Origine As String

#Region "Accesseurs"

    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des Clients
    'Retour : Rend une collection de Clients
    '=======================================================================
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNom As String = "", Optional ByVal strRS As String = "") As Collection
        Dim colReturn As Collection

        Persist.shared_connect()
        colReturn = ListeCLT(strCode, strNom, strRS)
        Persist.shared_disconnect()
        Return colReturn
    End Function

    'Identifiant du type de client
    Public Property idTypeClient() As Integer
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_idTypeClient
        End Get
        Set(ByVal Value As Integer)
            RaiseUpdated()
            m_idTypeClient = Value
        End Set
    End Property
    'Identifiant du type de client
    Public ReadOnly Property codeTypeClient() As String
        Get
            Dim oParam As Param
            oParam = New Param()
            oParam.load(idTypeClient)
            Return oParam.code
        End Get
    End Property

    'Libellé du type de client
    Public ReadOnly Property libTypeClient() As String
        Get
            Dim oParam As Param
            oParam = New Param()
            oParam.load(idTypeClient)
            Return oParam.code
        End Get

    End Property

    'code Tarif
    Public Property CodeTarif() As String
        Get
            Return m_CodeTarif
        End Get
        Set(ByVal Value As String)
            If Value <> CodeTarif Then
                RaiseUpdated()
                m_CodeTarif = Value
            End If
        End Set

    End Property
    ''' <summary>
    ''' Si le client est de type 'Intermédiaire' , indique l'origine des commande pour lequel il est un intermédiaire
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Origine() As String
        Get
            Return m_Origine
        End Get
        Set(ByVal Value As String)
            If Value <> Origine Then
                RaiseUpdated()
                m_Origine = Value
            End If
        End Set

    End Property
    Private ReadOnly Property colgPrecom() As ColEvent
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_oPreCommande.colLignes
        End Get
    End Property

    Public ReadOnly Property oPrecommande() As preCommande
        Get
            Return m_oPreCommande
        End Get
    End Property
    Public Shared Function createandload(ByVal pid As Long) As Client
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim objClient As Client
        Dim bReturn As Boolean
        objClient = New Client
        Try
            If pid <> 0 Then
                bReturn = objClient.load(pid)
                If Not bReturn Then
                    setError("Client.createAndLoad", getErreur())
                End If

            End If
        Catch ex As Exception
            setError("Client.createAndLoad", ex.ToString)
        End Try
        Debug.Assert(objClient.id = pid, "Client " & pid & " non chargée")
        Return objClient
    End Function 'Createanload
    ''' <summary>
    ''' Constructeur pour Chargement par la clé
    ''' </summary>
    ''' <param name="pCode"> Code Fournisseur</param>
    ''' <returns>Objet Fournisseur ou null</returns>
    ''' <remarks></remarks>
    Public Shared Function createandload(ByVal pCode As String) As Client
        Dim oClt As Client
        Dim bReturn As Boolean
        Dim nId As Integer
        oClt = New Client
        Try
            If Not String.IsNullOrEmpty(pCode) Then
                nId = Client.getCLTIDByKey(pCode)
                If nId <> -1 Then
                    bReturn = oClt.load(nId)
                    If Not bReturn Then
                        setError("Client.createAndLoad", getErreur())
                        oClt = Nothing
                    End If
                Else
                    setError("Client.createAndLoad", "No ID for " & pCode)
                    oClt = Nothing
                End If
            End If
        Catch ex As Exception
            setError("Client.createAndLoad", ex.ToString)
            oClt = Nothing
        End Try
        Return oClt
    End Function 'Createanload

    ''' <summary>
    ''' Constructeur pour Chargement par la clé
    ''' </summary>
    ''' <param name="pCode"> Code Fournisseur</param>
    ''' <returns>Objet Fournisseur ou null</returns>
    ''' <remarks></remarks>
    Public Shared Function createandloadPrestashop(ByVal pIdPrestashop As String) As Client
        Dim oClt As Client
        Dim bReturn As Boolean
        Dim nId As Integer
        shared_connect()
        oClt = New Client
        Try
            If Not String.IsNullOrEmpty(pIdPrestashop) Then
                nId = Client.getCLTIDByPrestashopId(pIdPrestashop)
                If nId <> -1 Then
                    bReturn = oClt.load(nId)
                    If Not bReturn Then
                        setError("Client.createandloadPrestashop", getErreur())
                        oClt = Nothing
                    End If
                Else
                    setError("Client.createandloadPrestashop", "No ID for " & pIdPrestashop)
                    oClt = Nothing
                End If
            End If
        Catch ex As Exception
            setError("Client.createandloadPrestashop", ex.ToString)
            oClt = Nothing
        End Try
        shared_disconnect()
        Return oClt
    End Function 'CreateanloadPrestashop
#End Region
    '=======================================================================
    'Fonction : savePrecommande
    'Description : Sauvegarde des lignes de precommandes
    'Détails    : 
    'Retour : 
    '=======================================================================
    Public Function savePrecommande() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(Not m_oPreCommande Is Nothing, "Precommande Existante")
        Persist.shared_connect()
        bReturn = m_oPreCommande.save()
        Persist.shared_disconnect()
        Return bReturn
    End Function 'savePrecommande
    '=======================================================================
    'Fonction : Load
    'Description : Chargement de l'objet en base
    'Détails    : Appelle LoadFRN (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        shared_connect()
        If pid <> 0 Then
            m_id = pid
        End If
        bReturn = loadCLT()
        m_oPreCommande = New preCommande(Me)
        m_oPreCommande.load(m_id)
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            m_bNew = False
            setUpdatedFalse()
            m_bDeleted = False
        End If
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : LoadLight
    'Description : Chargement du Résumé de  l'objet en base
    'Détails    : Appelle LoadFRN (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Friend Overrides Function LoadLight() As Boolean
        Debug.Assert(m_id <> 0, "L'id doit être initialisé")
        Dim bReturn As Boolean
        shared_connect()
        bReturn = loadCLTLight()
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            m_bNew = False
            setUpdatedFalse()
            m_bDeleted = False
        End If
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : LoadPreCommande
    'Description : Chargement de la precommande
    'Détails    : Appelle Loadprecommande (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Public Function LoadPreCommande() As Boolean
        Debug.Assert(m_id <> 0, "Le Client doit être sauvegardé au Préalable")
        Dim bReturn As Boolean
        bReturn = True
        shared_connect()
        bReturn = m_oPreCommande.load(m_id)
        shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : delete
    'Description : suppression de l'objet en base
    'Détails    : Appelle deleteFRN (de Persist)
    '               DELETEpRECOMMANDE
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        bReturn = deletePRECOMMANDE()
        If bReturn Then
            bReturn = deleteCLT()
            m_oPreCommande = Nothing
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
            If existeCommandeClient() Then
                bReturn = False
            End If
            shared_disconnect()
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete

    'Fonction : Insert
    'Description : Insert de l'objet en base
    'Détails    : Appelle InsertFRN (de Persist) 
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        bReturn = insertCLT()
        If (bReturn) Then
            m_oPreCommande.setClientId(Me)
        End If
        'If m_bcolLgPrecomLoaded Then
        'bReturn = savePrecommande()
        'End If
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : update
    'Description : Update de l'objet en base
    'Détails    : Appelle UpdateFRN (de Persist) 
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        bReturn = updateCLT()
        'If m_bcolLgPrecomLoaded Then
        'bReturn = savePrecommande()
        'End If
        Return bReturn
    End Function

    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        If (bReturn) Then
            If (Not m_oPreCommande Is Nothing) Then
                m_oPreCommande.save()
            End If
        End If
        shared_disconnect()
        Return bReturn
    End Function

    '=======================================================================
    'Fonction : Equals
    'Description : Comparaison d'objet
    'Détails    : Compare chaque membre de l'objet 
    'Retour : Rend Vrai si l'objet passé en paramétre est équivalent
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objclt As Client
        Dim bReturn As Boolean
        Try
            bReturn = True
            objclt = CType(obj, Client)

            bReturn = MyBase.Equals(obj)
            If bReturn Then
                bReturn = bReturn And (objclt.idTypeClient.Equals(Me.idTypeClient))
                bReturn = bReturn And (objclt.CodeTarif.Equals(Me.CodeTarif))
            End If

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function ' Equals

    '=======================================================================
    'Fonction : new
    'Description : Création de l'objet
    'Détails    : 
    'Retour : 
    '=======================================================================

    Public Sub New(ByVal pcode As String, ByVal pnom As String)
        MyBase.New(pcode, pnom)
        m_typedonnee = vncEnums.vncTypeDonnee.CLIENT
        Debug.Assert(Param.typeclientdefaut.defaut, "Pas de Type de Client par Defaut")
        m_idTypeClient = Param.typeclientdefaut.id
        m_libTypeClient = Param.typeclientdefaut.valeur
        m_oPreCommande = New preCommande(Me)
        m_CodeTarif = "A"
    End Sub 'New
    Friend Sub New()
        MyBase.New("", "")
        m_typedonnee = vncEnums.vncTypeDonnee.CLIENT
        Debug.Assert(Param.typeclientdefaut.defaut, "Pas de Type de Client par Defaut")
        m_idTypeClient = Param.typeclientdefaut.id
        m_libTypeClient = Param.typeclientdefaut.valeur
        m_oPreCommande = New preCommande(Me)
        m_CodeTarif = "A"
    End Sub 'New

    '=======================================================================
    'Fonction : toString
    'Description : Rend L'objet sous forme de chaine
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "CLT : [" & MyBase.toString() & "]" & idTypeClient & "," & libTypeClient
    End Function
    '=======================================================================
    'Fonction : AjouteLgPrecomm
    'Description : Ajoute une ligne LgPrecom au client parametre = Produit
    'Retour : Boolean
    '=======================================================================
    Public Function ajouteLgPrecom(ByVal objProduit As Produit, Optional ByVal qteHab As Double = 0, Optional ByVal qteDern As Double = 0, Optional ByVal pPrixU As Double = 0, Optional ByVal pDateDernCom As Date = DATE_DEFAUT, Optional ByVal prefDernCommande As String = "") As lgPrecomm
        Dim oLgPrecom As lgPrecomm
        oLgPrecom = m_oPreCommande.ajouteLgPrecom(objProduit, qteHab, qteDern, pPrixU, pDateDernCom, prefDernCommande)
        Return oLgPrecom
    End Function 'AjouteLgPrecomm
    '=======================================================================
    'Fonction : AjouteLgPrecomm
    'Description : Ajoute une ligne LgPrecom au client parametres (idProduit, CodeProduit,LibelleProduit)
    'Retour : LgPrecomm
    '=======================================================================
    Public Function ajouteLgPrecom(ByVal idProduit As Integer, ByVal codeProduit As String, ByVal libProduit As String, Optional ByVal qteHab As Decimal = 0, Optional ByVal qteDern As Decimal = 0, Optional ByVal pPrixU As Double = 0, Optional ByVal pDateDernCom As Date = DATE_DEFAUT, Optional ByVal prefDernCommande As String = "") As lgPrecomm
        Debug.Assert(idProduit <> 0, "ID Produit <> 0")

        Dim objLgPrecom As lgPrecomm
        objLgPrecom = m_oPreCommande.ajouteLgPrecom(idProduit, codeProduit, libProduit, qteHab, qteDern, pPrixU, pDateDernCom, prefDernCommande)
        Return objLgPrecom
    End Function 'AjouteLgPrecomm
    Public Function ajouteLgPrecom() As lgPrecomm

        Dim objLgPrecom As lgPrecomm
        objLgPrecom = m_oPreCommande.ajouteLgPrecom()
        Return objLgPrecom
    End Function 'AjouteLgPrecomm
    '=======================================================================
    'Fonction : SupprimeLgPrecomm
    'Description : Supprime une ligne de precommande au client
    'Retour : Boolean
    '=======================================================================
    Public Function supprimeLgPrecom(ByVal nIndex As Integer) As Boolean
        Try
            m_oPreCommande.supprimeLgPrecom(nIndex)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function 'supprimeLgPrecomm


    '=======================================================================
    'Fonction : getLgPrecom
    'Description : Renvoie la ligne de precommande
    'Retour : Boolean
    '=======================================================================
    Public Function getLgPrecomByProductId(ByVal pidProduit As Integer) As lgPrecomm
        Dim objLgPrecomm As lgPrecomm

        Try
            objLgPrecomm = m_oPreCommande.getLgPrecom(pidProduit)
        Catch ex As Exception
            objLgPrecomm = Nothing
        End Try
        Return objLgPrecomm
    End Function 'getLgPrecomm
    '=======================================================================
    'Fonction : lgPrecomExists
    'Description : Renvoie Vrai si la ligne de precommande existe
    'Retour : Boolean
    '=======================================================================
    Public Function lgPrecomExists(ByVal pindex As Object) As Boolean
        Dim bReturn As Boolean

        Try
            bReturn = m_oPreCommande.colLignes.keyExists(pindex)
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'lgPrecommExists
    '=======================================================================
    'Fonction : getlgPrecomCount
    'Description : Renvoie Le nombre de ligne de precommande
    'Retour : integer
    '=======================================================================
    Public Function getlgPrecomCount() As Integer
        Dim nReturn As Integer

        Try
            nReturn = m_oPreCommande.colLignes.Count
        Catch ex As Exception
            nReturn = 0
        End Try
        Return nReturn
    End Function 'lgPrecommCount

    '=======================================================================
    'Fonction : reinitPrecommande
    'Description : Réinitialise la precommande d'un client
    'Retour : Boolean
    '=======================================================================
    Public Function reinitPrecommande() As Boolean
        Dim bReturn As Boolean
        Dim colCommande As Collection
        Dim objCommandeClient As CommandeClient

        Try
            bReturn = True
            'Suppression des lignes existantes
            m_oPreCommande.colLignes.clear()
            m_oPreCommande.save()
            'chargement de la liste des commandes du client
            colCommande = CommandeClient.getListe(strCode:="", strNomClient:=Me.rs, pOrigine:="", pEtat:=vncEtatCommande.vncRien)
            'pour chaque commande
            For Each objCommandeClient In colCommande
                If objCommandeClient.oTiers.id = m_id Then
                    objCommandeClient.load()
                    objCommandeClient.loadcolLignes()
                    updatePrecommande(objCommandeClient)
                    objCommandeClient.releasecolLignes()
                End If
            Next objCommandeClient
            colCommande = Nothing
        Catch ex As Exception
            bReturn = False
            setError("Client.reinitPrecommande", ex.ToString)
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'reinitPrecommande
    '=======================================================================
    'Fonction : updatePrecommande
    'Description : Met à jour la précommande du client avec la commande en cours
    'Retour : Boolean
    '=======================================================================
    Public Function updatePrecommande(ByVal objCommande As CommandeClient) As Boolean
        Dim bColaDecharger As Boolean
        Dim objLGCMD As LgCommande
        Dim objLGPRECOM As lgPrecomm
        Dim bReturn As Boolean

        Try
            LoadPreCommande()
            bColaDecharger = False
            If Not objCommande.bcolLignesLoaded Then
                objCommande.loadcolLignes()
                bColaDecharger = True
            End If
            For Each objLGCMD In objCommande.colLignes
                'Parcours des lignes de commande
                'Pour chaque ligne
                If lgPrecomExists(objLGCMD.oProduit.id) Then
                    'Recupération de la ligne de precommande
                    objLGPRECOM = getLgPrecomByProductId(objLGCMD.oProduit.id)
                    If objLGPRECOM.dateDerniereCommande <= objCommande.dateCommande Then
                        'Si la date de dernierecommande est inférieur ou egale à la date de la commande
                        'MAJ du Prix, de la quantite , de la date et de la reference de la commande
                        objLGPRECOM.prixU = objLGCMD.prixU
                        objLGPRECOM.qteDern = objLGCMD.qteCommande
                        objLGPRECOM.dateDerniereCommande = objCommande.dateCommande
                        objLGPRECOM.refDerniereCommande = objCommande.code
                    End If
                Else
                    ajouteLgPrecom(objLGCMD.oProduit.id, objLGCMD.oProduit.code, objLGCMD.oProduit.nom, objLGCMD.qteCommande, objLGCMD.qteCommande, objLGCMD.prixU, objCommande.dateCommande, objCommande.code)
                End If
            Next objLGCMD
            If bColaDecharger Then
                objCommande.releasecolLignes()
            End If
            bReturn = True
        Catch ex As Exception
            bReturn = False
            setError("Client.updatePrecommande", ex.ToString)
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'updatePrecommande
    ''' <summary>
    ''' Fixe le type de client à "intermédaire" et l'origine des commandes concernées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function setTypeIntermediaire(pOrigine As String) As Boolean
        Dim bReturn As Boolean
        Try

            For Each oParam As Param In Param.colTypeClient
                If oParam.valeur = "Intermediaire" Or oParam.valeur = "Intermédiaire" Then
                    idTypeClient = oParam.id
                    Origine = pOrigine

                End If
            Next

            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function
    ''' <summary>
    ''' Renvoie le client qui a été identifié comme intermédiaire pour une origine
    ''' </summary>
    ''' <param name="pOrigine"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getIntermediairePourUneOrigine(pOrigine As String) As Client
        Dim oReturn As Client = Nothing
        Try
            Dim idTypeClient As Integer
            For Each oParam As Param In Param.colTypeClient
                If oParam.code.ToUpper() = "INT" Then
                    idTypeClient = oParam.id
                End If
            Next

            Dim strSQL As String
            strSQL = "SELECT CLT_ID FROM client where clt_Origine = '" & pOrigine & "' and CLT_TYPE_id = " & idTypeClient
            Dim strResultat As String = Persist.executeSQLQuery(strSQL)

            If Not String.IsNullOrEmpty(strResultat) Then
                oReturn = Client.createandload(CLng(strResultat))
            End If

        Catch ex As Exception
            oReturn = Nothing
        End Try
        Return oReturn
    End Function
End Class
