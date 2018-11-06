Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Collections.Generic
Imports System.Data.OleDb


'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : SousCommande 
' Description : SousCommandeClient
'               Hérite de Commande, crée par CommandeClient.creationSousCommandes
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
'20/09/04 : Ajout du lien sur la Facture de Commission
'
'===================================================================================================================================Public MustInherit Class Persist
Public Class SousCommande
    Inherits Commande


#Region "membres"
    Protected m_idCommandeClient As Integer
    Protected m_refFactFournisseur As String
    Protected m_dateFactFournisseur As Date
    Protected m_totalHTFacture As Decimal
    Protected m_totalTTCFacture As Decimal
    Protected m_baseCommission As Decimal
    Protected m_tauxCommission As Decimal
    Protected m_montantCommission As Decimal
    Protected m_oFournisseur As Fournisseur
    Protected m_idFactCom As Integer
    Protected m_codeCommande As String
    Protected m_bExportInternet As Boolean
    Protected m_bExportQuadra As Boolean
#End Region

#Region "Accesseurs"
    Public Sub New(ByVal poCommandeClient As CommandeClient, ByVal poFournisseur As Fournisseur)
        MyBase.New(poCommandeClient.oTiers, New EtatSSCommandeGeneree)
        'Réaffectation du Tiers (pour HOBIVIN le client de la SousCommande, n'est pas forcement le client de la commande)

        Dim oClient As Client
        oClient = Client.createandload(poCommandeClient.oTiers.id)
        setTiers(oClient)

        m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE
        Debug.Assert(Not poCommandeClient Is Nothing, "Commandeclient non renseignée")
        m_idCommandeClient = poCommandeClient.id
        dateCommande = poCommandeClient.dateCommande
        m_refFactFournisseur = ""
        m_dateFactFournisseur = DATE_DEFAUT
        m_baseCommission = 0
        m_totalHTFacture = 0
        m_totalTTCFacture = 0
        m_baseCommission = 0
        m_tauxCommission = 0
        CommFacturation.comment = CStr(poCommandeClient.CommFacturation.comment)
        m_oFournisseur = poFournisseur
        m_dateLivraison = poCommandeClient.dateLivraison
        m_RefLivraison = poCommandeClient.refLivraison
        m_oTransporteur.AdresseLivraison.nom = poCommandeClient.oTransporteur.AdresseLivraison.nom
        m_oTransporteur.AdresseLivraison.rue1 = poCommandeClient.oTransporteur.AdresseLivraison.rue1
        m_oTransporteur.AdresseLivraison.rue2 = poCommandeClient.oTransporteur.AdresseLivraison.rue2
        m_oTransporteur.AdresseLivraison.cp = poCommandeClient.oTransporteur.AdresseLivraison.cp
        m_oTransporteur.AdresseLivraison.ville = poCommandeClient.oTransporteur.AdresseLivraison.ville
        m_oTransporteur.AdresseLivraison.tel = poCommandeClient.oTransporteur.AdresseLivraison.tel
        m_oTransporteur.AdresseLivraison.fax = poCommandeClient.oTransporteur.AdresseLivraison.fax
        m_oTransporteur.AdresseLivraison.port = poCommandeClient.oTransporteur.AdresseLivraison.port
        m_oTransporteur.AdresseLivraison.Email = poCommandeClient.oTransporteur.AdresseLivraison.Email
        m_idFactCom = 0
        m_codeCommande = poCommandeClient.code
        m_bExportInternet = False
        m_bExportQuadra = False
        majBooleenAlaFinDuNew()

        Origine = poCommandeClient.Origine

    End Sub

    Friend Sub New()
        'Contructeur utilisée uniquement dans Cette DLL (Persist)
        MyClass.New(New CommandeClient(New Client), New Fournisseur)
    End Sub

    Public Overloads Shared Function createandload(ByVal pid As Long) As SousCommande
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim objSousCommande As SousCommande
        objSousCommande = New SousCommande
        Try
            If pid <> 0 Then
                objSousCommande.load(pid)
            End If
        Catch ex As Exception
            setError("CreateSousCommande", ex.ToString)
        End Try
        Debug.Assert(objSousCommande.id = pid, "Sous Commande " & pid & " non chargée")
        Return objSousCommande
    End Function 'createandload
    Public Overloads Shared Function createandload(ByVal pstrCode As String) As SousCommande
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Debug.Assert(pstrCode <> "", "Le code n'est pas renseigné")
        Dim objSousCommande As SousCommande
        objSousCommande = New SousCommande
        Try
            objSousCommande.loadFromCode(pstrCode)
        Catch ex As Exception
            setError("CreateSousCommande", ex.ToString)
        End Try
        Debug.Assert(objSousCommande.id <> 0, "Sous Commande " & pstrCode & " non chargée")
        Return objSousCommande
    End Function 'createandload
    '=======================================================================
    'Fonction : Load
    'Description : Chargement d'un Element Persistant
    '           Appel DBLoad qui doit être redéfinis dans la classe de base
    'Retour : 
    '=======================================================================
    Public Function loadFromCode(ByVal pstrCode As String) As Boolean
        Dim bReturn As Boolean
        Dim bResumeOld As Boolean
        shared_connect()

        bResumeOld = m_bResume
        m_bResume = False
        bReturn = loadSCMD(pstrCode)
        If bReturn Then
            'Chargement des caractéristiques du Client
            bReturn = oTiers.load()
            Debug.Assert(bReturn, Tiers.getErreur())
            bReturn = oFournisseur.load()
            Debug.Assert(bReturn, Fournisseur.getErreur())
        End If
        'm_bcolLgLoaded = loadcolLignes()
        Return bReturn
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            resetBooleans()
        Else
            m_bResume = bResumeOld
        End If
        'Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' load

    Public Property idCommandeClient() As Integer
        Get
            Return m_idCommandeClient
        End Get
        Set(ByVal Value As Integer)
            If Not m_idCommandeClient.Equals(Value) Then
                m_idCommandeClient = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property refFactFournisseur() As String
        Get
            Return m_refFactFournisseur
        End Get
        Set(ByVal Value As String)
            If Not m_refFactFournisseur.Equals(Value) Then
                m_refFactFournisseur = Value
                RaiseUpdated()
            End If

        End Set
    End Property

    Public Property dateFactFournisseur() As Date
        Get
            Return m_dateFactFournisseur
        End Get
        Set(ByVal Value As Date)
            If Not m_dateFactFournisseur.Equals(Value) Then
                m_dateFactFournisseur = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property totalHTFacture() As Decimal
        Get
            Return m_totalHTFacture
        End Get
        Set(ByVal Value As Decimal)
            If Not m_totalHTFacture.Equals(Value) Then
                m_totalHTFacture = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    Public Property totalTTCFacture() As Decimal
        Get
            Return m_totalTTCFacture
        End Get
        Set(ByVal Value As Decimal)
            If Not m_totalTTCFacture.Equals(Value) Then
                m_totalTTCFacture = Value
                RaiseUpdated()
            End If

        End Set
    End Property

    Public Property baseCommission() As Decimal
        Get
            Return m_baseCommission
        End Get
        Set(ByVal Value As Decimal)
            If Not m_baseCommission.Equals(Value) Then
                m_baseCommission = Value
                RaiseUpdated()
            End If

        End Set
    End Property

    Public Property tauxCommission() As Decimal
        Get
            Return m_tauxCommission
        End Get
        Set(ByVal Value As Decimal)
            If Not m_tauxCommission.Equals(Value) Then
                m_tauxCommission = Value
                RaiseUpdated()
            End If

        End Set
    End Property

    Public Overloads Property MontantCommission() As Decimal
        Get
            Return m_montantCommission
        End Get
        Set(ByVal Value As Decimal)
            If Not m_montantCommission.Equals(Value) Then
                m_montantCommission = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    '20/10/08 suppression de la pahse de provisionnment
    'Public Property dateProvCommission() As Date
    '    Get
    '        Return m_dateProvCommission.ToShortDateString
    '    End Get
    '    Set(ByVal Value As Date)
    '        If Not m_dateProvCommission.ToShortDateString.Equals(Value.ToShortDateString) Then
    '            m_dateProvCommission = Value.ToShortDateString
    '            RaiseUpdated()
    '        End If
    '    End Set
    'End Property

    Public ReadOnly Property oFournisseur() As Fournisseur
        Get
            Return m_oFournisseur
        End Get
    End Property

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return code & " | " & oFournisseur.rs & " | " & dateCommande & " | " & Format(totalHT, "C") & " | " & etat.libelle
        End Get
    End Property
    Public ReadOnly Property shortResumeClient() As String
        Get
            Dim str As String
            str = code & " | " & oFournisseur.code & " | " & oTiers.rs & " | " & dateCommande & " | " & Format(totalHT, "C") & " | " & etat.libelle
            Return str
        End Get
    End Property
    Public ReadOnly Property shortResumeFournisseur() As String
        Get
            Dim str As String
            str = refFactFournisseur & " | " & dateFactFournisseur.ToShortDateString & " | " & oTiers.rs & " | " & Format(totalHT, "C") & "|" & code & " | " & dateCommande & " | " & oFournisseur.code & "|" & etat.libelle
            Return str
        End Get
    End Property
    Public ReadOnly Property shortResumeFactCom() As String
        Get
            Dim str As String
            str = refFactFournisseur & " | " & oFournisseur.code & " | " & oTiers.rs & " | " & Format(dateFactFournisseur, "d") & " | " & Format(MontantCommission, "C")
            Return str
        End Get
    End Property


    Public Property idFactCom() As Integer
        Get
            Return m_idFactCom
        End Get
        Set(ByVal Value As Integer)
            If Not m_idFactCom.Equals(Value) Then
                m_idFactCom = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    'Le code de la commande n'est pas mis ç jour dans la BD
    Public Property codeCommandeClient() As String
        Get
            Return m_codeCommande
        End Get
        Set(ByVal Value As String)
            m_codeCommande = Value
        End Set
    End Property

    'Conversion du tiers de la comande en client
    Public ReadOnly Property oClient() As Client
        Get
            Return CType(oTiers, Client)
        End Get
    End Property

    Public Property bExportInternet() As Boolean
        Get
            Return m_bExportInternet
        End Get
        Set(ByVal Value As Boolean)
            If Not m_bExportInternet.Equals(Value) Then
                m_bExportInternet = Value
                RaiseUpdated()
            End If

        End Set
    End Property
    Public Property bExportQuadra() As Boolean
        Get
            Return m_bExportQuadra
        End Get
        Set(ByVal Value As Boolean)
            If Not m_bExportQuadra.Equals(Value) Then
                m_bExportQuadra = Value
                RaiseUpdated()
            End If

        End Set
    End Property

    Public Overrides ReadOnly Property FournisseurRS() As String
        Get
            If m_oFournisseur IsNot Nothing Then
                Return m_oFournisseur.rs
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public Overrides ReadOnly Property FournisseurCode() As String
        Get
            If m_oFournisseur IsNot Nothing Then
                Return m_oFournisseur.code
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property Fournisseurid() As Integer
        Get
            If m_oFournisseur IsNot Nothing Then
                Return m_oFournisseur.id
            Else
                Return 0
            End If
        End Get
    End Property



#End Region
#Region "Fonction communes"
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "Sous Commande : " & MyBase.toString()
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As SousCommande
        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj

            bReturn = bReturn And (m_idCommandeClient.Equals(objCommande.idCommandeClient))
            bReturn = bReturn And (m_refFactFournisseur.Equals(objCommande.refFactFournisseur))
            bReturn = bReturn And (m_dateFactFournisseur.Equals(objCommande.dateFactFournisseur))
            bReturn = bReturn And (m_totalHTFacture.Equals(objCommande.totalHTFacture))
            bReturn = bReturn And (m_totalTTCFacture.Equals(objCommande.totalTTCFacture))
            bReturn = bReturn And (m_tauxCommission.Equals(objCommande.tauxCommission))
            bReturn = bReturn And (m_montantCommission.Equals(objCommande.MontantCommission))
            bReturn = bReturn And (m_bExportInternet.Equals(objCommande.bExportInternet))
            bReturn = bReturn And (m_bExportQuadra.Equals(objCommande.bExportQuadra))

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals
    Public Overrides Function genereMvtStock() As Boolean
        Return False
    End Function 'genereMvtStock

    Public Overrides Function supprimeMvtStock() As Boolean
        Return False
    End Function 'supprimeMvtStock
    Friend Overrides Function setNewcode() As Boolean
        Dim ncode As Integer
        Dim breturn As Boolean

        shared_connect()
        ncode = getNumeroSousCommandeClient()
        shared_disconnect()
        If ncode = -1 Then
            breturn = False
        Else
            code = m_codeCommande & "-" & CStr(ncode)
            breturn = True
        End If
        Return breturn
    End Function 'setnewCode

#End Region

#Region "Interface Persist"
    '================================================================================
    'Fonction: DBLoad 
    'Description : Chargement de l'objet courant
    '================================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean

        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        Debug.Assert(id <> 0, "idSousCommande <> 0")
        bReturn = loadSCMD()
        If bReturn Then
            'Chargement des caractéristiques du Client
            bReturn = oTiers.load()
            If Not bReturn Then
                Debug.Assert(bReturn, Tiers.getErreur())
            End If
            bReturn = oFournisseur.load()
            Debug.Assert(bReturn, Fournisseur.getErreur())
        End If
        'm_bcolLgLoaded = loadcolLignes()
        Return bReturn
    End Function 'dbLoad
    '================================================================================
    'Fonction: Delete
    'Description : Suppression de la sous commande
    '               la suppression de la sous-commande n'entraine pas la supression des lignes
    '================================================================================

    Friend Overrides Function delete() As Boolean
        Debug.Assert(id <> 0, "idSCommande <> 0")
        Dim bReturn As Boolean
        Dim objLg As LgCommande
        '#857 : Les lignes de sous commandes sont maintenant indépendante des lignes de commandes
        bReturn = deletecolLgSCMD()  'suppression des lignes de sous-commandes
        Debug.Assert(bReturn, getErreur)
        colLignes.clear()
        m_bcolLgLoaded = False
        m_bColLgInsertorDelete = False

        'Suppression de la sousCommande
        bReturn = deleteSCMD()

        'Suppression du Lien entre la commande et les sous-commandes
        If bReturn Then
            Dim oCmdCLT As CommandeClient
            oCmdCLT = CommandeClient.createandload(Me.idCommandeClient)
            oCmdCLT.loadcolLignes()
            For Each oLg As LgCommande In oCmdCLT.colLignes
                If oLg.idSCmd = Me.id Then
                    oLg.idSCmd = 0
                End If
            Next
            oCmdCLT.savecolLignes()
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
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete
    '================================================================================
    'Fonction: insert
    'Description : Création de la sous commande en base
    '               La Création d'une sous-commande entraine la Mise à jour des lignes
    '================================================================================

    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas renseigné")
        Debug.Assert(Not oFournisseur Is Nothing, "Le Fournisseur n'est pas renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(oFournisseur.id <> 0, "Le fournisseur n'est pas sauvegardé")
        Debug.Assert(id = 0, "idSousCommande = 0")
        Debug.Assert(idCommandeClient <> 0, "L'id de la commande Client doit être renseigné")

        Dim objLg As LgCommande
        Dim bReturn As Boolean

        If code = "" Then
            setNewcode()
        End If
        'Création de la sous-commande en base (récupération de l'id)
        bReturn = insertSCMD()
        For Each objLg In colLignes
            objLg.idSCmd = Me.id
            objLg.bLigneEclatee = True
        Next
        'Mise à jour des lignes de Sous-commandes
        If bReturn Then
            bReturn = INSERTcolLGCMD()
        End If

        Return bReturn

    End Function 'insert

    Friend Overrides Function update() As Boolean
        '================================================================================
        'Fonction: update
        'Description : Mise à jour de la sous commande en base
        '               La Mise à jour d'une sous-commande entraine la Mise à jour des lignes
        '================================================================================
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas renseigné")
        Debug.Assert(Not oFournisseur Is Nothing, "Le Fournisseur n'est pas renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(oFournisseur.id <> 0, "Le fournisseur n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idCommande <> 0")

        Dim bReturn As Boolean
        Dim objLG As LgCommande

        'Mise à jour de la sous-commande
        bReturn = updateSCMD()
        Debug.Assert(bReturn, "sous-commande.update" & getErreur())
        If bReturn Then
            For Each objLG In colLignes
                objLG.idSCmd = Me.id
                objLG.bLigneEclatee = True
            Next
            bReturn = UPDATEcolLGCMD()
        End If
        Return bReturn
    End Function 'Update
#End Region

#Region "Méthodes"
    Public Overrides Function changeEtat(ByVal p_action As vncActionEtatCommande) As Boolean
        MyBase.changeEtat(p_action)
        If etat.codeEtat = vncEnums.vncEtatCommande.vncSCMDRapprochee And m_bExportInternet Then
            etat = EtatCommande.createEtat(vncEtatCommande.vncSCMDRapprocheeInt)
        End If
        RaiseUpdated()
    End Function
    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des Sous Commande
    'Retour : Rend une collection de Sous commande
    '=======================================================================
    Public Shared Function getListe(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)
        shared_connect()
        colReturn = ListeSCMD(pddeb, pdfin, pCodeFourn, pEtat)
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListe
    Public Shared Function getListeTransmises(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "") As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)
        Dim colTemp As List(Of SousCommande)
        Dim objSCMD As SousCommande

        colReturn = New List(Of SousCommande)
        shared_connect()
        colTemp = ListeSCMD(pddeb, pdfin, pCodeFourn, vncEnums.vncEtatCommande.vncSCMDtransmiseFax)
        If Not colTemp Is Nothing Then
            For Each objSCMD In colTemp
                colReturn.Add(objSCMD)
            Next
        End If
        colTemp = ListeSCMD(pddeb, pdfin, pCodeFourn, vncEnums.vncEtatCommande.vncSCMDExporteeInt)
        If Not colTemp Is Nothing Then
            For Each objSCMD In colTemp
                colReturn.Add(objSCMD)
            Next
        End If
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListeTransmises
    '========================================================================
    '/// <summary>Rend une liste de sous commande A provisonner = Rapprochée + RapprochéeInternet</summary>
    '/// <param name="pddeb">date de debut</param>
    '/// <param name="pdfin">date de fin</param>
    '/// <returns>une collection</returns>
    '========================================================================
    Public Shared Function getListeAProvisionner(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "") As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)
        Dim colTemp As List(Of SousCommande)
        Dim objSCMD As SousCommande

        colReturn = New List(Of SousCommande)
        shared_connect()
        colTemp = ListeSCMD(pddeb, pdfin, pCodeFourn, vncEnums.vncEtatCommande.vncSCMDRapprochee)
        If Not colTemp Is Nothing Then
            For Each objSCMD In colTemp
                colReturn.Add(objSCMD)
            Next
        End If
        colTemp = ListeSCMD(pddeb, pdfin, pCodeFourn, vncEnums.vncEtatCommande.vncSCMDRapprocheeInt)
        If Not colTemp Is Nothing Then
            For Each objSCMD In colTemp
                colReturn.Add(objSCMD)
            Next
        End If
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListeAProvisioner

    '========================================================================
    '/// <summary>Rend une liste de sous commande A Facturer = Rapprochée + RapprochéeInternet + Provisionnée</summary>
    '/// <param name="pddeb">date de debut</param>
    '/// <param name="pdfin">date de fin</param>
    '/// <returns>une collection</returns>
    '========================================================================
    Public Shared Function getListeAFacturer(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT, Optional ByVal pCodeFourn As String = "") As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)

        shared_connect()
        colReturn = ListeSCMDAFacturerCom(pddeb, pdfin, pCodeFourn)

        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListeAFacturer
    '========================================================================
    '/// <summary>Rend une liste de sous commande exportées</summary>
    '/// <param name="pddeb">date de debut</param>
    '/// <param name="pdfin">date de fin</param>
    '/// <returns>une collection</returns>
    '========================================================================
    Public Shared Function getListeExportee(Optional ByVal pddeb As Date = DATE_DEFAUT, Optional ByVal pdfin As Date = DATE_DEFAUT) As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)
        shared_connect()
        colReturn = ListeSCMD(pddeb, pdfin, , vncEnums.vncEtatCommande.vncSCMDExporteeInt)
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListeExportee
    '========================================================================
    '/// <summary>Rend une liste de sous commande exportées</summary>
    '/// <param name="pddeb">date de debut</param>
    '/// <param name="pdfin">date de fin</param>
    '/// <returns>une collection</returns>
    '========================================================================
    ''' <summary>
    ''' Rend une liste de sous commande à Exporter par internet 
    ''' </summary>
    ''' <param name="pTypeExport">Type d'export Fournisseur (1=ExportInternet, 2 = ExportQuadra)</param>
    ''' <param name="pOrigine">Origine de la Commande</param>
    ''' <param name="pddeb">Date De Debut</param>
    ''' <param name="pdfin">Date de fin </param>
    ''' <param name="pStrCodeFourn">Code Fournisseur (Facultatif)</param>
    ''' <returns></returns>
    Public Shared Function getListeAExporterInternet(pOrigine As vncOrigineCmd, ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pStrCodeFourn As String = "") As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)
        shared_connect()
        colReturn = ListeSCMDAExporter(pOrigine, pddeb, pdfin, pStrCodeFourn, vncTypeExportScmd.vncExportInternet, vncTypeExportScmd.vncExportInternet)
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListeAExporterInternet
    ''' <summary>
    ''' Rend une liste de sous commande à Exporter QUADRA 
    ''' </summary>
    ''' <param name="pTypeExportFRN">Type d'export Fournisseur (1=ExportInternet, 2 = ExportQuadra)</param>
    ''' <param name="pOrigine">Origine de la Commande</param>
    ''' <param name="pddeb">Date De Debut</param>
    ''' <param name="pdfin">Date de fin </param>
    ''' <param name="pStrCodeFourn">Code Fournisseur (Facultatif)</param>
    ''' <returns></returns>
    Public Shared Function getListeAExporterQuadra(pTypeExportFRN As vncTypeExportScmd, pOrigine As vncOrigineCmd, ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pStrCodeFourn As String = "") As List(Of SousCommande)
        Dim colReturn As List(Of SousCommande)
        shared_connect()
        colReturn = ListeSCMDAExporter(pOrigine, pddeb, pdfin, pStrCodeFourn, pTypeExportFRN, vncTypeExportScmd.vncExportQuadra)
        Debug.Assert(Not colReturn Is Nothing, getErreur())
        shared_disconnect()
        Return colReturn
    End Function 'getListeAExporterQuadra

    ''' <summary>
    ''' Valider l'export quadra : L'état de la sousCommande ne change pas,
    ''' le Boolean bExportquadra est mis à jour
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function ValiderExportQuadra() As Boolean
        Debug.Assert(bcolLignesLoaded, "Les Lignes doivent être chargées au préalable")
        Dim objLgCommande As LgCommande
        Dim bReturn As Boolean
        bReturn = False

        Try
            'Me.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
            'Me.bExportInternet = True
            'Quantité Facturée = Quantité Livrée
            For Each objLgCommande In colLignes
                objLgCommande.qteFact = objLgCommande.qteLiv
            Next objLgCommande
            If Me.Origine = Dossier.HOBIVIN Then
                If Me.oFournisseur.bExportInternet = vncTypeExportScmd.vncExportQuadra Then
                    'C'est un Founisseur 'Hobivin' sur une commande Hobivin
                    'changement d'état de la sous Commandes
                    'Les commandes exportées vers Quadra sont déclarer facturée car Quadra ne fait pas de rapprochement
                    Me.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDExportInternet)
                    Me.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
                    Me.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)
                    Me.refFactFournisseur = "QUADRAFACT"

                End If
            End If
            'changement d'état de la sous Commandes
            'Les commandes exportées vers Quadra sont déclarer facturée car Quadra ne fait pas de rapprochement
            'Me.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDRapprocher)
            'Me.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)
            Me.bExportQuadra = True
            bReturn = True
        Catch ex As Exception
            setError("SousCommande.ValiderExportQuadra ERR" & ex.Message)
            bReturn = False
        End Try
        Return bReturn

    End Function

    '========================================================================
    '/// <summary>Genere le bon à Facturer dans un fichier PDF</summary>
    '/// <param name="strPathToReport">Emplacement des états Crystal Report</param>
    '/// <param name="strPDFFileName">Nom du fichier à générer</param>
    '/// <returns>True if OK</returns>
    '========================================================================
    Public Function genererPDF(ByVal strPathToReport As String, ByVal strPDFFileName As String) As Boolean
        Debug.Assert(Trim(strPathToReport) <> "", "PathToReport non initialisé")
        Debug.Assert(Trim(strPDFFileName) <> "", "strPDFFileName non initialisé")
        Dim bReturn As Boolean
        Dim oReport As ReportDocument
        Dim diskOpts As New CrystalDecisions.Shared.DiskFileDestinationOptions
        Try
            oReport = genererReport(strPathToReport)
            If Not oReport Is Nothing Then
                oReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                oReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                diskOpts.DiskFileName = strPDFFileName
                oReport.ExportOptions.DestinationOptions = diskOpts
                oReport.Export()
                oReport.Close()
                bReturn = True
            Else
                bReturn = False
            End If
        Catch ex As Exception
            setError("SousCommande.genererPDF", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function 'genererPDF
    '========================================================================
    ' Fonction : genererReport
    'Description  : Charge et generer le rapport de sous-commande
    '========================================================================
    Public Function genererReport(ByVal strPathtoReport As String) As ReportDocument
        Dim bReturn As Boolean
        Dim oReport As ReportDocument = Nothing

        Try
            bReturn = True
            oReport = New ReportDocument
            oReport.Load(strPathtoReport & "crBonFournisseur.rpt")
            Persist.setReportConnection(oReport)
            oReport.Refresh()
            oReport.SetParameterValue("IDSCMD", id)

        Catch ex As Exception
            bReturn = False
            oReport = Nothing
            setError("SousCommande.genererReport", ex.ToString())
        End Try
        Return oReport
    End Function 'genererReport
    '===============================================================================================
    'Fonction : faxer
    'Description : génère le rapport de la sous commande et le faxe au fournisseur par defaut au numéro de fax de l'adresse de facturation
    '===============================================================================================
    Public Function faxer(ByVal pstrPathtoReport As String, ByVal strNomInterlocteur As String, ByVal strTelInterlocuteur As String, ByVal strSubject As String, ByVal strNotes As String, ByVal bSendCoverPage As Boolean, ByVal strfilename As String, Optional ByVal pstrNumFax As String = "") As Boolean
        Dim bReturn As Boolean
        'Dim objFax As clsFax
        'Dim strNumFax As String
        'Dim diskOpts As New CrystalDecisions.Shared.DiskFileDestinationOptions
        'Dim discreteVal As New ParameterDiscreteValue
        'Dim nJobId As Integer
        'Dim oReport As ReportDocument

        'Try
        '    If Trim(pstrNumFax) = "" Then
        '        strNumFax = oFournisseur.AdresseFacturation.fax
        '    Else
        '        strNumFax = pstrNumFax
        '    End If
        '    oReport = genererReport(pstrPathtoReport)

        '    If Not oReport Is Nothing Then
        '        oReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        '        oReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        '        diskOpts.DiskFileName = strfilename
        '        oReport.ExportOptions.DestinationOptions = diskOpts
        '        oReport.Export()
        '        objFax = New clsFax
        '        nJobId = m_id
        '        nJobId = objFax.sendFax(strNomInterlocteur, strTelInterlocuteur, strSubject, strNotes, bSendCoverPage, diskOpts.DiskFileName, strNumFax, oFournisseur)
        '        oReport.Close()
        '        'WaitFiveSeconds()
        '        objFax = Nothing
        '        If nJobId > 0 Then
        '            bReturn = True
        '        Else
        '            bReturn = False
        '        End If
        '    End If

        'Catch ex As Exception
        '    bReturn = False
        '    setError("SousCommande.faxer", ex.ToString)
        'End Try

        Return bReturn
    End Function 'faxer
    '=======================================================================
    'Fonction : AjouteLigne()
    'Description : Ajoute une ligne sur une SousCommande
    'Détails    :  
    'Retour : une ligne de commande ou nothing si l'ajout échoue
    '=======================================================================
    Public Overloads Function AjouteLigne(ByVal pobjLgCmd As LgCommande, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not pobjLgCmd Is Nothing)
        Debug.Assert(pobjLgCmd.idCmd <> 0, "L'ID de la commande n'est pas renseigné dans la ligne ")
        Debug.Assert(pobjLgCmd.id <> 0, "la ligne n'est pas sauvegardé ")
        Debug.Assert(pobjLgCmd.oProduit.idFournisseur = oFournisseur.id, "Founisseur Différents")

        Dim oReturn As LgCommande

        oReturn = pobjLgCmd.Clone()
        oReturn.bLigneEclatee = True
        '        oReturn = pobjLgCmd
        oReturn.idSCmd = Me.id
        'Ajout du tiers dans la Ligne de commmande (utile pour touver le Tx de commission)
        oReturn.idTiers = oTiers.id

        Try
            If p_bCalculPrix Then
                oReturn.calculPrixTotal()
            End If
            m_colLignes.Add(oReturn, CStr(oReturn.num))
            If p_bCalculPrix Then
                calculPrixTotal()
            End If
            m_bColLgInsertorDelete = True
            m_bcolLgLoaded = True
        Catch ex As Exception
            setError("SousCommande.AjouteLigne", "Ajout de ligne impossible key = " & pobjLgCmd.num)
            oReturn = Nothing
        End Try
        '        Debug.Assert(m_ocolLignes.Count = n + 1, "Le nombre d'élement dans la collection est incrémenté de 1")
        Return oReturn
    End Function 'AjouteLigne
    ''' <summary>
    ''' Export d'une Sous commande au format CSV pour être exportée vers le site internet
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function toCSV() As String
        Debug.Assert(bcolLignesLoaded, "Les Lignes doivent être chargées au préalable")
        Dim strResult1Line As String
        Dim strResult As String
        Dim objCommande As CommandeClient
        Dim objLgCommande As LgCommande
        Dim objProduit As Produit

        If idCommandeClient <> 0 Then
            objCommande = CommandeClient.createandload(idCommandeClient)
        Else
            'Normalement ceci ne sert pour les tests
            objCommande = New CommandeClient(Me.oClient)
        End If

        strResult = ""
        For Each objLgCommande In colLignes
            strResult1Line = ""
            objProduit = Produit.createandload(objLgCommande.oProduit.id) 'Chargement du produit
            strResult1Line = strResult1Line & Trim(Me.id) & ";"
            strResult1Line = strResult1Line & Trim(Me.code) & ";"
            strResult1Line = strResult1Line & Trim(Me.codeCommandeClient) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.dateCommande, "ddMMyyyy")) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.code) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.nom) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.rs) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.nom) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.rue1) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.rue2) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.cp) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.ville) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.tel) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.fax) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.port) & ";"
            strResult1Line = strResult1Line & Trim(Me.oFournisseur.AdresseFacturation.Email) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.nom) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.rue1) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.rue2) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.cp) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.ville) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.tel) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.fax) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.port) & ";"
            strResult1Line = strResult1Line & Trim(Me.oTransporteur.AdresseLivraison.Email) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.code) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.nom) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.rs) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.nom) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.rue1) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.rue2) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.cp) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.ville) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.tel) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.fax) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.port) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseLivraison.Email) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.nom) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.rue1) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.rue2) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.cp) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.ville) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.tel) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.fax) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.port) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.AdresseFacturation.Email) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.libTypeClient) & ";"
            strResult1Line = strResult1Line & Trim(objCommande.typeTransport) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.libModeReglement) & ";"
            strResult1Line = strResult1Line & Trim(objCommande.refLivraison) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.banque) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.rib1) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.rib2) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.rib3) & ";"
            strResult1Line = strResult1Line & Trim(Me.oClient.rib4) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.dateLivraison, "ddMMyyyy")) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.dateEnlevement, "ddMMyyyy")) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.totalHT, "n")) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.totalTTC, "n")) & ";"
            strResult1Line = strResult1Line & Trim(Me.CommFacturation.comment) & ";"
            strResult1Line = strResult1Line & Trim(objLgCommande.num) & ";"
            strResult1Line = strResult1Line & Trim(objLgCommande.qteLiv) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.code) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.nom) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.libCouleur) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.millesime) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.libConditionnement) & ";"
            strResult1Line = strResult1Line & Trim(objProduit.libContenant) & ";"
            strResult1Line = strResult1Line & Trim(objLgCommande.prixU) & ";"
            strResult1Line = strResult1Line & Trim(Me.refFactFournisseur) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.dateFactFournisseur, "ddMMyyyy")) & ";"
            strResult1Line = strResult1Line & totalHTFacture & ";"
            strResult1Line = strResult1Line & totalTTCFacture & ";"
            'Exportation des Taux et Montants de commissions
            strResult1Line = strResult1Line & tauxCommission & ";"

            strResult1Line = Replace(strResult1Line, vbCrLf, "--")
            strResult1Line = Replace(strResult1Line, vbCr, "-")
            strResult1Line = Replace(strResult1Line, vbLf, "-")
            strResult1Line = Replace(strResult1Line, vbNullChar, "-")
            strResult1Line = Replace(strResult1Line, vbTab, "-")
            strResult1Line = Replace(strResult1Line, vbBack, "-")
            strResult = strResult & strResult1Line & vbCrLf
        Next objLgCommande
        Return strResult

    End Function 'ToCSV





    '/// <summary>Create a CSV string for simulate internet export
    '/// export only Id and datas fom fournisseur
    '///</summary>
    '/// <value>CSV string</value>
    '/// <exceptions cref="ArgumentNullException">Thrown when the specified value is Nothing (C#, VC++: null)<exceptions>
    Public Function FTO_toCSVFromInternet() As String
        Debug.Assert(bcolLignesLoaded, "Les Lignes doivent être chargées au préalable")
        Dim strResult1Line As String
        Dim strResult As String
        Dim objLgCommande As LgCommande


        strResult = ""
        For Each objLgCommande In colLignes
            strResult1Line = ""
            strResult1Line = strResult1Line & Trim(Me.id) & ";"
            'Dans ce mode le détail de la sous commande n'est pas exporté
            strResult1Line = strResult1Line & Trim(Me.refFactFournisseur) & ";"
            strResult1Line = strResult1Line & Trim(Format(Me.dateFactFournisseur, "ddMMyyyy")) & ";"
            strResult1Line = strResult1Line & totalHTFacture & ";"
            strResult1Line = strResult1Line & totalTTCFacture & ";"

            strResult1Line = Replace(strResult1Line, vbCrLf, "--")
            strResult1Line = Replace(strResult1Line, vbCr, "-")
            strResult1Line = Replace(strResult1Line, vbLf, "-")
            strResult1Line = Replace(strResult1Line, vbNullChar, "-")
            strResult1Line = Replace(strResult1Line, vbTab, "-")
            strResult1Line = Replace(strResult1Line, vbBack, "-")
            strResult = strResult & strResult1Line & vbCrLf
        Next objLgCommande
        Return strResult

    End Function 'toCSVForTestFromInternet

    Public Overrides Function calculPrixTotal() As Boolean
        '========================================================================
        ' CalculPrixTotal : Calcul du prix de lacommande, du poids et des qte de colis
        '       Utilise le fonction CalculPtixTotal de la classe Commande pour calculer les prixs
        '========================================================================
        Dim bReturn As Boolean
        Try
            bReturn = MyBase.calculPrixTotal()
            For Each oLg As LgCommande In m_colLignes

                'Calcul de la commssion pour Chaque Ligne
                If (Me.etat.codeEtat = vncEtatCommande.vncEnCoursSaisie Or Me.etat.codeEtat = vncEtatCommande.vncValidee) Then
                    oLg.CalculCommission(Origine, CalculCommQte.CALCUL_COMMISSION_QTE_CMDE)
                Else
                    oLg.CalculCommission(Origine, CalculCommQte.CALCUL_COMMISSION_QTE_LIVREE)
                End If
            Next oLg
        Catch ex As Exception
            setError("CommandeClient.CalculPrixTotal", ex.ToString)
            bReturn = False
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'calculPrixTotal


    ''' <summary>
    ''' Calcul de la commission pour la sousCommande
    ''' </summary>
    ''' <param name="pBaseComm">base de commission (HT Calculé ou HT Facturé)</param>
    ''' <returns></returns>
    ''' <remarks>Taux de commission = Taux de la première ligne</remarks>
    Public Function calcCommisionstandard(ByVal pBaseComm As vncEnums.CalculCommScmd) As Boolean
        Dim bReturn As Boolean
        Dim objLG As LgCommande
        Try
            If colLignes.Count = 0 Then
                bReturn = loadcolLignes()
            End If

            Select Case pBaseComm
                Case CalculCommScmd.CALCUL_COMMISSION_HT_CALCULE
                    baseCommission = totalHT
                Case CalculCommScmd.CALCUL_COMMISSION_HT_FACTURE
                    baseCommission = totalHTFacture
                Case CalculCommScmd.CALCUL_COMMISSION_BASE_COMM
                    baseCommission = baseCommission
            End Select
            If tauxCommission = 0 Then
                'Taux de commssion par defaut
                tauxCommission = Param.getConstante("CST_TX_COMMISSION")
                'S'il y a des lignes ils prend le taux de commission associée à la ligne 
                'Ce taux a été calculé au moment de la saisie de commande.
                If colLignes.Count > 0 Then
                    objLG = colLignes(1)
                    If (objLG IsNot Nothing) Then
                        tauxCommission = objLG.TxComm
                    End If
                End If
            End If
            MontantCommission = Math.Round(baseCommission * (tauxCommission / 100), 2)

            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'calcCommisionstandard
#End Region

    Public Overrides Sub Exporter(ByVal pfileNAme As String)
        ' no export for SousCommande
    End Sub

    Public Overrides Function ControleMvtStock() As Microsoft.VisualBasic.Collection
        Return Nothing
    End Function

    Public Function DuppliqueInfosCmd() As Boolean
        Dim bReturn As Boolean
        Try
            totalHTFacture = totalHT
            totalTTCFacture = totalTTC
            dateFactFournisseur = Now().ToShortDateString()
            baseCommission = totalHT
            calcCommisionstandard(CalculCommScmd.CALCUL_COMMISSION_HT_FACTURE)
            bReturn = True
        Catch ex As Exception
            setError(ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function

#Region "Surchargée Methode pour l'export Quadra"
    Public Overrides Function getCodeCommande() As String
        Return codeCommandeClient
    End Function
    Public ReadOnly Property ClientRS() As String
        Get
            Return oClient.rs
        End Get
    End Property

#End Region
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        If m_bcolLgLoaded Then
            bReturn = bReturn And savecolLignes()
        End If
        shared_disconnect()
        Return bReturn
    End Function ' Save
    '=======================================================================
    'Fonction : savecolLignes()
    'Description : Sauvegarde la collection des lignes
    '               En fonction du paramètre bDeleteInsert
    '                   Suppression des lignes pour reinsertion (CommandeClient)
    '                   Update des lignes
    'Détails    :  
    'Retour : Boolean
    '=======================================================================
    Public Overrides Function savecolLignes() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(Not m_colLignes Is Nothing, "m_col <> nothing")
        Debug.Assert(m_bcolLgLoaded, "La collection  doit être chargée au préalable")
        Debug.Assert(m_id <> 0, "La commande  doit être sauvegardée au préalable")
        'En mode commandeClient ilfaut supprimer les lignes avant de les recréer pour gérer les suppressions et ajouts de lignes.
        If m_bColLgInsertorDelete Then
            'On Supprime la collection avant de la recréer (Commande)
            bReturn = deletecolLgSCMD()
            If bReturn Then
                bReturn = INSERTcolLGSCMD()
            End If
            m_bColLgInsertorDelete = False
        Else
            'On Met à jour la collection (SousCommande)
            bReturn = UPDATEcolLGSCMD()
        End If
        Debug.Assert(bReturn, "Commande.savecolLignes:" & getErreur())
        Return bReturn
    End Function
#Region "Methode Persist"
    Protected Function INSERTcolLGSCMD() As Boolean
        Dim objCMD As Commande
        Dim bReturn As Boolean
        Dim objLgCMD As LgCommande

        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Le Client doit être Sauvegardé")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE _
                        , "Objet de type SousCommande requis")
        objCMD = Me
        Debug.Assert(Not objCMD.colLignes Is Nothing, "ColLignes is Nothing")


        Dim sqlString As String = "INSERT INTO LIGNE_COMMANDE (" & _
                                    "LGCM_NUM," & _
                                    "LGCM_CODE," & _
                                    "LGCM_CMD_ID," & _
                                    "LGCM_SCMD_ID," & _
                                    "LGCM_PRD_ID," & _
                                    "LGCM_QTE_COMMANDE," & _
                                    "LGCM_QTE_LIV," & _
                                    "LGCM_QTE_FACT," & _
                                    "LGCM_PRIX_UNITAIRE," & _
                                    "LGCM_PRIX_HT," & _
                                    "LGCM_PRIX_TTC," & _
                                    "LGCM_BGRATUIT," & _
                                    "LGCM_BECLATEE, " & _
                                    "LGCM_POIDS, " & _
                                    "LGCM_QTE_COLIS, " & _
                                    "LGCM_TXCOMM, " & _
                                    "LGCM_MTCOMM, " & _
                                    "LGCM_BA_ID " & _
                                    ") VALUES ( " & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "? ," & _
                                    "?" & _
                                  " )"
        Dim objCommand As OleDbCommand
        Dim objCommand2 As OleDbCommand
        Dim objRS As OleDbDataReader = Nothing
        '        Dim objParam As OleDbParameter

        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction
        'Commande de lecture de l'id
        objCommand2 = New OleDbCommand("SELECT MAX(LGCM_ID) FROM LIGNE_COMMANDE", m_dbconn.Connection)
        objCommand2.Transaction = m_dbconn.transaction


        bReturn = True
        Dim oParam1 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        Dim oParam2 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.VarChar)
        Dim oParamIDCmd As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        oParamIDCmd.IsNullable = True
        Dim oParamIDSCMD As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        oParamIDSCMD.IsNullable = True
        Dim oParam5 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        Dim oParam6 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam7 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam8 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam9 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam10 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam11 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam12 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Boolean)
        Dim oParam13 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Boolean)
        Dim oParam14 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParam15 As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Double)
        Dim oParamTxComm As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Decimal)
        Dim oParamMtComm As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Decimal)
        Dim oParamIDBa As OleDbParameter = objCommand.Parameters.AddWithValue("?", OleDbType.Integer)
        oParamIDBa.IsNullable = True
        For Each objLgCMD In objCMD.colLignes
            Try
                If m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE Then
                    objLgCMD.idBA = 0
                    objLgCMD.idCmd = 0
                    objLgCMD.idSCmd = id
                End If

                oParam1.Value = objLgCMD.num
                oParam2.Value = (objCMD.code & objLgCMD.num)
                If objLgCMD.idCmd = 0 Or objLgCMD.idCmd = -1 Then
                    oParamIDCmd.Value = DBNull.Value
                Else
                    oParamIDCmd.Value = (objLgCMD.idCmd)
                End If
                If objLgCMD.idSCmd = 0 Or objLgCMD.idSCmd = -1 Then
                    oParamIDSCMD.Value = DBNull.Value
                Else
                    oParamIDSCMD.Value = (objLgCMD.idSCmd)
                End If
                oParam5.Value = (objLgCMD.oProduit.id)
                oParam6.Value = (CDbl(objLgCMD.qteCommande))
                oParam7.Value = (CDbl(objLgCMD.qteLiv))
                oParam8.Value = (CDbl(objLgCMD.qteFact))
                oParam9.Value = (CDbl(objLgCMD.prixU))
                oParam10.Value = (CDbl(objLgCMD.prixHT))
                oParam11.Value = (CDbl(objLgCMD.prixTTC))
                oParam12.Value = (objLgCMD.bGratuit)
                oParam13.Value = (objLgCMD.bLigneEclatee)
                oParam14.Value = (CDbl(objLgCMD.poids))
                oParam15.Value = (CDbl(objLgCMD.qteColis))
                oParamTxComm.Value = objLgCMD.TxComm
                oParamMtComm.Value = objLgCMD.MtComm
                If objLgCMD.idBA = 0 Or objLgCMD.idBA = -1 Then
                    oParamIDBa.Value = DBNull.Value
                Else
                    oParamIDBa.Value = (objLgCMD.idBA)
                End If
                objCommand.ExecuteNonQuery()
                objRS = objCommand2.ExecuteReader
                If objRS.Read() Then
                    objLgCMD.setid(objRS.GetInt32(0))
                End If
                cleanErreur()
                objRS.Close()
            Catch ex As Exception
                setError("InsertcolLGSCmd", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next
        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' insertcolLGSCMD
    Protected Function UPDATEcolLGSCMD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "La Commande.id  <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE, "Objet de type SousCommande requis")


        Dim sqlString As String = "UPDATE LIGNE_COMMANDE SET " & _
                                    " LGCM_CMD_ID = ? ," & _
                                    " LGCM_SCMD_ID= ? ," & _
                                    " LGCM_BA_ID= ? ," & _
                                    " LGCM_NUM= ? ," & _
                                    " LGCM_PRD_ID = ? ," & _
                                    " LGCM_QTE_COMMANDE = ? ," & _
                                    " LGCM_QTE_LIV = ? ," & _
                                    " LGCM_QTE_FACT = ? ," & _
                                    " LGCM_PRIX_UNITAIRE = ? ," & _
                                    " LGCM_PRIX_HT = ? ," & _
                                    " LGCM_PRIX_TTC= ? ," & _
                                    " LGCM_POIDS = ? , " & _
                                    " LGCM_QTE_COLIS = ? , " & _
                                    " LGCM_TXCOMM = ? , " & _
                                    " LGCM_MTCOMM = ? , " & _
                                    " LGCM_BGRATUIT = ? , " & _
                                    " LGCM_BECLATEE = ?" & _
                                    " WHERE " & _
                                    " LGCM_ID = ?"
        Dim objCommand As OleDbCommand
        Dim objCMD As Commande
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean
        Dim objlgCMD As LgCommande

        objCMD = CType(Me, Commande)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString
        m_dbconn.BeginTransaction()
        objCommand.Transaction = m_dbconn.transaction



        bReturn = True
        Dim oParamIDCMD As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParamIDSCMD As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParamIDBA As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParam4 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParam5 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        Dim oParam6 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam7 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam8 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam9 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam10 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam11 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam12 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        Dim oParam13 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        'Taux de commission
        Dim oParamTxComm As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)
        'Montant de commssion
        Dim oParamMtcomm As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Decimal)

        Dim oParam14 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Boolean)
        Dim oParam15 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Boolean)
        Dim oParam16 As System.Data.OleDb.OleDbParameter = objCommand.Parameters.Add("?", OleDbType.Integer)
        For Each objlgCMD In objCMD.colLignes
            Try
                If (objlgCMD.idCmd = 0) Then
                    oParamIDCMD.Value = DBNull.Value
                Else
                    oParamIDCMD.Value = (objlgCMD.idCmd)
                End If
                If objlgCMD.idSCmd = 0 Then
                    oParamIDSCMD.Value = DBNull.Value
                Else
                    oParamIDSCMD.Value = (objlgCMD.idSCmd)
                End If
                If (objlgCMD.idBA = 0) Then
                    oParamIDBA.Value = DBNull.Value
                Else
                    oParamIDBA.Value = (objlgCMD.idBA)
                End If

                oParam4.Value = (objlgCMD.num)
                oParam5.Value = (objlgCMD.oProduit.id)
                oParam6.Value = (objlgCMD.qteCommande)
                oParam7.Value = (objlgCMD.qteLiv)
                oParam8.Value = (objlgCMD.qteFact)
                oParam9.Value = (objlgCMD.prixU)
                oParam10.Value = (objlgCMD.prixHT)
                oParam11.Value = (objlgCMD.prixTTC)
                oParam12.Value = (objlgCMD.poids)
                oParam13.Value = (objlgCMD.qteColis)
                oParam14.Value = (objlgCMD.bGratuit)
                oParam15.Value = (objlgCMD.bLigneEclatee)
                oParam16.Value = (objlgCMD.id)
                oParamTxComm.Value = (objlgCMD.TxComm)
                oParamMtcomm.Value = (objlgCMD.MtComm)
                objCommand.ExecuteNonQuery()
                'objRS.Close()
            Catch ex As Exception
                setError("UpdatecolLGSCMD", ex.ToString())
                bReturn = False
                Exit For
            End Try
        Next objlgCMD

        If bReturn Then
            m_dbconn.transaction.Commit()
        Else
            m_dbconn.transaction.Rollback()
        End If

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function ' UPDATEcolLGCMD
    '==========================================================================
    'Methode : deletecolLGCMD
    'Description : Suppression des lignes d'une commande
    '               la clé de destrcution est choisie en fonction de la classe de l'objet
    Protected Function deletecolLgSCMD() As Boolean
        Debug.Assert(shared_isConnected(), "La database doit être ouverte")
        Debug.Assert(m_id <> 0, "Id <> 0")
        Debug.Assert(m_typedonnee = vncEnums.vncTypeDonnee.SSCOMMANDE _
                        , "Objet de type SousCommande requis")


        Dim sqlString As String = "DELETE FROM LIGNE_COMMANDE "
        Dim strClauseWhereSCMD As String = " WHERE LIGNE_COMMANDE.LGCM_SCMD_ID = ? "
        Dim objCMD As Commande
        Dim objCommand As OleDbCommand
        '        Dim objParam As OleDbParameter
        Dim bReturn As Boolean

        'Clause where en Fonction du type de l'objet courant
        sqlString = sqlString & strClauseWhereSCMD

        objCMD = CType(Me, Commande)
        objCommand = New OleDbCommand
        objCommand.Connection = m_dbconn.Connection
        objCommand.CommandText = sqlString



        CreateParameterP_ID(objCommand)
        Try
            objCommand.ExecuteNonQuery()

            bReturn = True

        Catch ex As Exception
            setError("DeletecolLGSCMD", ex.ToString())
            bReturn = False
        End Try
        Debug.Assert(bReturn, "DeletecolLGSCMD" & getErreur())
        Return bReturn
    End Function 'DeletecolLGSCMD

#End Region

End Class
