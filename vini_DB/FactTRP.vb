Imports System.configuration
Public Class FactTRP
    Inherits Facture

    Private m_dateStat As Date
    Private m_periode As String
    Private m_montantTaxes As Decimal
    Private m_montantReglement As Decimal
    Private m_refReglement As String
    Private m_dateReglement As Date
    '    Private m_idModeRegelement As Long

#Region "Accesseurs"
    Public Sub New(ByVal poClient As Client)
        MyBase.New(poClient, EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactTRPGeneree))
        m_periode = ""
        m_montantTaxes = 0
        m_montantReglement = 0
        m_refReglement = ""
        m_dateReglement = DATE_DEFAUT
        m_typedonnee = vncEnums.vncTypeDonnee.FACTTRP
        idModeReglement = poClient.idModeReglement1 'utilisation du mode de reglement du Tiers
    End Sub
    Public Shared Function createandload(ByVal pid As Long) As FactTRP
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactTRP
        obj = New FactTRP
        Try
            If pid <> 0 Then
                obj.load(pid)
            End If
        Catch ex As Exception
            setError("FactTRP.CreateandLoad", ex.ToString)
        End Try
        Debug.Assert(obj.id = pid, "FactTRP " & pid & " non chargée")
        Return obj
    End Function 'createandload
    Public Shared Function createandload(ByVal pCode As String) As FactTRP
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactTRP = Nothing
        Dim oCol As Collection
        Try
            obj = New FactTRP()
            oCol = getListe(pCode)
            If (oCol.Count = 1) Then
                obj = oCol(1)
            End If

        Catch ex As Exception
            setError("CreateFactTRP", ex.Message)
        End Try
        Debug.Assert(obj.code = pCode, "FactTRP " & pCode & " non chargée")
        Return obj
    End Function 'createandload
    Friend Sub New()
        Me.New(New Client)
    End Sub
    Public Property dateReglement() As Date
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_dateReglement.ToShortDateString
        End Get
        Set(ByVal Value As Date)
            If Value.ToShortDateString <> m_dateReglement.ToShortDateString Then
                RaiseUpdated()
                m_dateReglement = Value.ToShortDateString
            End If
        End Set
    End Property

    Public Property dateStatistique() As Date
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_dateStat.ToShortDateString
        End Get
        Set(ByVal Value As Date)
            If Value.ToShortDateString <> m_dateStat.ToShortDateString Then
                RaiseUpdated()
                m_dateStat = Value.ToShortDateString
            End If
        End Set
    End Property
    Public Property periode() As String
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_periode
        End Get
        Set(ByVal Value As String)
            If Value <> m_periode Then
                RaiseUpdated()
                m_periode = Value
            End If
        End Set

    End Property
    Public Property refReglement() As String
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_refReglement
        End Get
        Set(ByVal Value As String)
            If Value <> m_refReglement Then
                RaiseUpdated()
                m_refReglement = Value
            End If
        End Set

    End Property
    Public Property montantReglement() As Decimal
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_montantReglement
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_montantReglement Then
                RaiseUpdated()
                m_montantReglement = Value
            End If
        End Set
    End Property

    Public Property montantTaxes() As Decimal
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_montantTaxes
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_montantTaxes Then
                RaiseUpdated()
                m_montantTaxes = Value
            End If
        End Set

    End Property

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return code & " | " & oTiers.rs & " | " & dateCommande.ToShortDateString & " | " & Format(totalTTC, "C") & " | " & etat.libelle
        End Get
    End Property

#End Region
#Region "Méthode de racine"

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "(" & code & "," & dateCommande & "," & periode & "," & oTiers.code & "," & m_TotalHT & "," & m_TotalTTC & ")"
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As FactTRP

        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj
            bReturn = bReturn And (m_periode.Equals(objCommande.periode))

            bReturn = bReturn And (m_montantReglement.Equals(objCommande.montantReglement))
            bReturn = bReturn And (m_dateReglement.Equals(objCommande.dateReglement))
            bReturn = bReturn And (m_refReglement.Equals(objCommande.refReglement))
            bReturn = bReturn And (m_IDModeReglement.Equals(objCommande.idModeReglement))
            bReturn = bReturn And (m_montantTaxes.Equals(objCommande.montantTaxes))
            Return bReturn

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals
#End Region
#Region "Interface Persist"
    Public Overrides Function checkForDelete() As Boolean
        Return True
    End Function
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean = False
        Try
            If pid <> 0 Then
                m_id = pid
            End If

            Debug.Assert(id <> 0, "idFactTRP <> 0")
            bReturn = LoadFACTTRP()
            If id <> 0 Then
                'Chargement des caractéristiques du Client
                Try
                    oTiers.load()
                Catch ex As Exception
                    bReturn = False
                    setError("oTiers.load", Tiers.getErreur())
                End Try
                Debug.Assert(bReturn, Tiers.getErreur())
                'Chargement des lignes de Facture
                'm_bcolLgLoaded = loadcolLignes()
            End If
            Return bReturn
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactTRP.DBLoad: " & getErreur())
        Return bReturn
    End Function 'DBLoad

    Friend Overrides Function delete() As Boolean

        Dim bReturn As Boolean = False
        Try
            bReturn = True
            If bcolLignesLoaded Then
                bReturn = deletecolLgTRP()
            End If
            If bReturn Then
                bReturn = deleteRefFACTTRP
            End If
            If bReturn Then
                bReturn = deleteFACTTRP
            End If
            If bReturn Then
                bReturn = deleteReglements()
            End If
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactTrp.delete: " & getErreur())
        Return bReturn

    End Function 'delete

    Public Overrides Function genereMvtStock() As Boolean
        Return False

    End Function

    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(id = 0, "id= 0")

        Dim bReturn As Boolean = False

        Try
            bReturn = False
            If setNewcode() Then
                bReturn = insertFACTTRP()
            End If
            If bcolLignesLoaded Then
                bReturn = savecolLignes()
            End If
        Catch ex As Exception
            bReturn = False
            setError("FactTRP.insert", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactTRP.insert: " & getErreur())
        Return bReturn

    End Function 'insert

    Public Overrides Function supprimeMvtStock() As Boolean
        Return False
    End Function

    Friend Overrides Function update() As Boolean
        '================================================================================
        'Fonction: update
        'Description : Mise à jour de la Facture de Transport
        '================================================================================
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idFact <> 0")
        Dim bReturn As Boolean = False
        Try
            'Mise à jour de la sous-commande
            bReturn = UpdateFACTTRP()
            If bcolLignesLoaded Then
                bReturn = savecolLignes()
            End If
        Catch ex As Exception
            bReturn = False
            setError("FactTRP.Update", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactTRP.Update: " & getErreur())
        Return bReturn

    End Function

    Public Overrides Function loadcolLignes() As Boolean
        Debug.Assert(m_id <> 0, "La Facture doit être sauvegardé au Préalable")
        Dim bReturn As Boolean = False
        shared_connect()
        If m_bcolLgLoaded Then
            m_colLignes.clear()
        End If
        bReturn = LoadcolLgFactTRP()

        Debug.Assert(bReturn, ColEvent.getErreur())
        m_bcolLgLoaded = bReturn ' Les lignes sont chargées
        m_bColLgInsertorDelete = False ' Mais ne sont pas Mise à jour
        shared_disconnect()
        Return bReturn
    End Function 'loadColLignes
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
            'On Supprime la collection des lignes avant de la recréer 
            '1° Suppression de la reference à la facture de transport dans la table Commande
            bReturn = deleteRefFACTTRP()
            If bReturn Then
                bReturn = deletecolLgTRP()
                If bReturn Then
                    bReturn = INSERTcolLGTRP()
                End If
                m_bColLgInsertorDelete = False
            End If
        Else
            'On Met à jour la collection (SousCommande)
            bReturn = UPDATEcolLGTRP()
        End If
        Debug.Assert(bReturn, "FactTRP.savecolLignes:" & getErreur())
        Return bReturn
    End Function

    Public Function purge() As Boolean
        '================================================================================
        ' Fonction : purge 
        ' Description : Supression de la facture ET de tous ses élements dépendants
        '================================================================================
        Dim bReturn As Boolean = False
        'Try
        '    bReturn = True
        '    If Not etat.codeEtat = vncEnums.vncEtatCommande.vncFactComRapprochee Then
        '        bReturn = False
        '        Exit Function
        '    End If
        '    'Chargement des lignes de factures de comm (Sous-Commandes)
        '    If Not m_bcolLgLoaded Then
        '        bReturn = loadcolLignes()
        '    End If
        '    If bReturn Then
        '        For Each objLgCMD In m_colLignes
        '            objSCMD = SousCommande.createandload(objLgCMD.idSCmd)
        '            Debug.Assert(Not objSCMD Is Nothing)
        '            'Suppression de la sous-commande
        '            objSCMD.bDeleted = True
        '            bReturn = objSCMD.Save()
        '            'La Commande n'est pas supprimée par cette purge
        '            'Voir purge des commandes
        '        Next
        '    End If
        '    If bReturn Then
        '        m_colLignes = New ColEvent   'Pour Empêcher le traitrement des Sous- Commandes
        '        Me.bDeleted = True
        '        bReturn = Save()
        '    End If

        'Catch ex As Exception
        '    setError("FactCom.purge", ex.ToString())
        '    bReturn = False
        'End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'purge

#End Region
#Region "Méthodes de commande"
    Friend Overrides Function setNewcode() As Boolean
        Dim str As String
        Dim ncode As Integer
        Dim breturn As Boolean

        str = ""
        shared_connect()
        ncode = GetNumeroFactureTransport()
        shared_disconnect()
        If ncode = -1 Then
            breturn = False
        Else
            code = str & CStr(ncode)
            breturn = True
        End If
        Return breturn
    End Function 'setnewCode
    '=======================================================================
    'Fonction : getNextNumLg()
    'Description : Rend le prochain Numéro de Ligne
    'Détails    :  
    'Retour : une Numéro de ligne
    '=======================================================================
    Public Overrides Function getNextNumLg() As Integer
        Dim oLg As LgFactTRP
        Dim num As Integer = 0
        Dim bOk As Boolean
        'Nombre d'élement dans la collection + 10
        num = ((m_colLignes.Count + 1) * 10)

        'Recherche du premier Element suivant de libre
        While Not bOk
            Try
                oLg = m_colLignes(CStr(num))
                num = num + 1
            Catch ex As Exception
                bOk = True
            End Try
        End While

        Return num
    End Function ' getNextNumLg
    Public Overloads Function AjouteLigne(ByVal pobjLgCMD As LgCommande, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "Méthode désactivée dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "Méthode désactivée dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overrides Function calculPrixTotal() As Boolean
        Dim oLg As LgFactTRP = Nothing
        Dim oLgGO As LgFactTRP = Nothing
        Dim nHT As Decimal
        Dim nHTCmd As Decimal
        Dim nTTC As Decimal
        Dim objParam As Param

        nHT = 0
        nHTCmd = 0
        nTTC = 0
        For Each oLg In m_colLignes
            nHT = nHT + oLg.prixHT
            'Pour le total Taxes GO on ne prend en compte que les lignes <> GO
            If oLg.num <> CST_LGFACTTRP_NUM_GO Then
                nHTCmd = nHTCmd + oLg.prixHT
            End If
        Next oLg


        If Param.getConstante("CST_TRP_TXGAZOLE") <> 0 Then
            'Recherche de la ligne GO
            If colLignes.keyExists(CST_LGFACTTRP_NUM_GO) Then
                oLgGO = colLignes(CST_LGFACTTRP_NUM_GO)
            Else
                oLgGO = New LgFactTRP
                oLgGO.num = CST_LGFACTTRP_NUM_GO
                oLgGO.nomTransporteur = CST_LGFACTTRP_LIB_GO & Param.getConstante("CST_TRP_TXGAZOLE") & "%"
                AjouteLigneFactTRP(oLgGO, False)
            End If
            oLgGO.prixHT = nHTCmd * (Param.getConstante("CST_TRP_TXGAZOLE") / 100)
        End If

        nHT = nHTCmd + montantTaxes
        If Not oLgGO Is Nothing Then
            nHT = nHT + oLgGO.prixHT
        End If

        objParam = New Param
        objParam.load(m_idParamTVA)

        nTTC = Math.Round(nHT * (1 + (objParam.valeur / 100)), 2)
        'On passe par les accesseurs pour lever l'event Updated s'il y a lieu
        totalHT = nHT
        totalTTC = nTTC
        Return True
    End Function 'calculPrixTotal

    Public Overrides Function changeEtat(ByVal p_action As vncActionEtatCommande) As Boolean
        Debug.Assert(p_action >= vncActionEtatCommande.vncActionMinFactTRP And p_action <= vncActionEtatCommande.vncActionMaxFactTRP)
        Dim bReturn As Boolean = False
        Try
            If p_action >= vncActionEtatCommande.vncActionMinFactTRP And p_action <= vncActionEtatCommande.vncActionMaxFactTRP Then
                MyBase.changeEtat(p_action)
                RaiseUpdated()
                bReturn = True
            Else
                setError("FactTRP.changeEtat", "Code Action incorrect :" + p_action)
                bReturn = False
            End If

        Catch ex As Exception
            setError("FactTRP.changeEtat", ex.ToString)
            bReturn = False
        End Try

        Return bReturn
    End Function 'ChangeEtat

    Public Overrides Function ControleMvtStock() As Microsoft.VisualBasic.Collection
        Return Nothing
    End Function


#End Region
    Public Shared Function createFactTRPs(ByVal pcolCMD As Collection, Optional ByVal pdateFact As Date = DATE_DEFAUT, Optional ByVal pdateStat As Date = DATE_DEFAUT, Optional ByVal pPeriode As String = "") As ColEvent

        '======================================================================================
        'Function : createFactTRPs
        'description : création des factures de Transport à partir d'une collection de commande
        '               Pour chaque Client Différent une facture de transport est crée et les 
        '           Commande de ce client lui sont attribué
        '======================================================================================
        Debug.Assert(Not pcolCMD Is Nothing)
        Dim ocolReturn As ColEvent
        Dim objCMD As CommandeClient
        Dim objFactTRP As FactTRP

        Try
            'Traitement des paramètres
            If pdateFact = DATE_DEFAUT Then
                pdateFact = Now()
            End If
            If pdateStat = DATE_DEFAUT Then
                pdateStat = pdateFact
            End If
            If pPeriode = "" Then
                pPeriode = CStr(pdateStat.ToString("d"))
            End If

            ocolReturn = New ColEvent
            'Parcours de toutes les Commandes de la collection
            For Each objCMD In pcolCMD
                'A-t-on déjà crée une facture de Transport pour ce Client
                If ocolReturn.keyExists(objCMD.oTiers.code) Then
                    'Une Facture à déjà été créée pour ce Client
                    objFactTRP = ocolReturn(objCMD.oTiers.code)
                Else
                    'Il n'y a pas de facture pour ce client
                    'Création de la facture de transport
                    objFactTRP = New FactTRP(objCMD.oTiers)
                    objFactTRP.dateFacture = pdateFact
                    objFactTRP.dateStatistique = pdateStat
                    objFactTRP.periode = pPeriode
                    objFactTRP.montantTaxes = 0
                    'Ajout de la facture à la collection 
                    ocolReturn.Add(objFactTRP, objFactTRP.oTiers.code)
                End If
                'Ajout de la taxe d'enregistrement autant de fois qu'il y a de lignes
                objFactTRP.montantTaxes = objFactTRP.montantTaxes + Param.getConstante("CST_TAXES_TRP")
                'Ajout de la Commande dans la facture
                objFactTRP.AjouteLigneFactTRP(objCMD, True)
            Next
        Catch ex As Exception
            ocolReturn = Nothing
            setError("FactTRP.createFactTRP", ex.ToString())
        End Try

        Debug.Assert(Not ocolReturn Is Nothing, getErreur())
        Return ocolReturn

    End Function 'createFactsTRP
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNomClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = ListeFACTTRP(strCode, strNomClient, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListe" & getErreur())
        shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListeNonReglee(Optional ByVal strCode As String = "", Optional ByVal strNomClient As String = "") As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeFACTTRPNonReglee(strCode, strNomClient)
        Debug.Assert(Not colReturn Is Nothing, "FactTRP.getListeNonReglee" & getErreur())
        shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListe(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pCodeClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeFACTTRPEtat(pddeb, pdfin, pCodeClient, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function
    Public Function AjouteLigneFactTRP(ByVal p_oCMD As CommandeClient, Optional ByVal pCalcul As Boolean = True) As LgFactTRP
        ''=======================================================================
        ''Fonction : AjouteLigneFactTRP()
        ''Description : Créé une ligne de Facture à partir de la commande et l'ajoute à la collection via AjouteLigne
        ''Détails    :   Appelle la Fonction AjoutLigne a
        ''Retour : une ligne de commande ou nothing si l'ajout échoue
        ''=======================================================================
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not p_oCMD Is Nothing)
        Debug.Assert(p_oCMD.id <> 0, "La commande doit être enregistrée")
        Debug.Assert(p_oCMD.oTiers.id = Me.oTiers.id, "Les identifiants de Clients doivent être égaux")
        Debug.Assert(p_oCMD.bFactTransport = True, "La commande doit accepter les factures de transports")
        Debug.Assert(p_oCMD.etat.codeEtat = vncEnums.vncEtatCommande.vncLivree Or _
                    p_oCMD.etat.codeEtat = vncEnums.vncEtatCommande.vncEclatee Or _
                    p_oCMD.etat.codeEtat = vncEnums.vncEtatCommande.vncRapprochee _
                    , "La commande doit être livrée")

        Dim oLg As LgFactTRP
        'Dim strNum As String

        Try
            oLg = New LgFactTRP
            oLg.num = CStr(Me.getNextNumLg())
            oLg.idFactTRP = Me.id 'L'id de la facture de trabsport peut -être à Zéro
            'On controle la non existence de la commande dan sla facture de transport
            If Not colLignes.keyExists(oLg.num()) Then
                'Dupplication des informations de la commande
                oLg.dupliqueinfosCommande(p_oCMD)
                'Affectation de l'id de la facture à la commande
                'Attention l'objet n'est pas sauvegardé en base 
                'La sauvegarde de la ligne de transport entrainera la MAJ de la commande
                p_oCMD.idFactTransport = Me.id 'L'id de la facture de trabsport peut -être à Zéro

                colLignes.Add(oLg, oLg.num)
                If pCalcul Then
                    calculPrixTotal()
                End If
                m_bcolLgLoaded = True
                m_bColLgInsertorDelete = True
            Else
                setError("FactTRP.AjouteLigneFactTRP", " : Ligne existante" + oLg.num)
                oLg = Nothing
            End If
        Catch ex As Exception
            setError("AjouteLigneFactTRP", ex.ToString)
            oLg = Nothing
        End Try
        Debug.Assert(Not oLg Is Nothing, getErreur())
        Return oLg
    End Function 'AjouteLigneFactCom

    Public Function AjouteLigneFactTRP(ByVal p_oLgFactTRP As LgFactTRP, Optional ByVal pCalcul As Boolean = True) As LgFactTRP
        ''=======================================================================
        ''Fonction : AjouteLigneFactTRP()
        ''Description : Ajoute une ligne à la facture (permet d'ajouter une ligne qui n'a pas de commande)
        ''Retour : une ligne de commande ou nothing si l'ajout échoue
        ''=======================================================================
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not p_oLgFactTRP Is Nothing)
        'Debug.Assert(Me.id <> 0, "La Facture doit être enregistrée")

        Dim oLg As LgFactTRP
        'Dim strNum As String

        If p_oLgFactTRP.num <> CST_LGFACTTRP_NUM_GO Then
            p_oLgFactTRP.num = CStr(getNextNumLg())
        End If
        Try

            'On n'utilise pas de clé dans cette méthode car les lignes qui n'ont pas de commande rattachés sont incontrolables
            '            If Not colLignes.keyExists(p_oLgFactTRP.key()) Then
            p_oLgFactTRP.idFactTRP = Me.id
            colLignes.Add(p_oLgFactTRP, p_oLgFactTRP.num)
            If pCalcul Then
                calculPrixTotal()
            End If
            m_bcolLgLoaded = True
            m_bColLgInsertorDelete = True
            oLg = p_oLgFactTRP
        Catch ex As Exception
            setError("AjouteLigneFactTRP", ex.ToString)
            oLg = Nothing
        End Try
        Debug.Assert(Not oLg Is Nothing, getErreur())
        Return oLg
    End Function 'AjouteLigneFactTRP
    '=======================================================================
    'Fonction : supprimeLigne()
    'Description : Supprime une ligne sur une commande
    'Détails    :  Le Numéro passé est le numéro de ligne et non l'indice dans la collection
    '
    'Retour : une ligne de commande ou nothing si l'ajout échoue
    '=======================================================================
    Public Overrides Function supprimeLigne(ByVal strNumLigne As String, Optional ByVal p_bCalculPrix As Boolean = True) As Boolean
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(m_bcolLgLoaded, "La collection des lignes doit être chargée")
        Dim bReturn As Boolean
        Dim objLg As LgFactTRP

        Try
            objLg = m_colLignes(strNumLigne)
            If objLg.num = strNumLigne Then
                m_colLignes.Remove(strNumLigne)
                If p_bCalculPrix Then
                    calculPrixTotal()
                End If
                bReturn = True
                m_bColLgInsertorDelete = True
            Else
                setError("FactTRP.SupprimeLigne(" & strNumLigne & ")", "Ligne inconnue")
                bReturn = False
            End If
        Catch ex As Exception
            setError("FactTRP.SupprimeLigne(" & strNumLigne & ")", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SupprimeLigneTRP

    ''========================================================================
    ''fonction : exporter
    ''DEscription : Exporte la facture de transport dans un format MAXIMA
    ''DEtails    : 1facture =1 Ligne avec le prix Total
    ''
    ''Retour : Boolean
    ''=========================================================================
    'Public Function exporter(ByVal strFile As String) As Boolean

    '    Debug.Assert(strFile <> "", "Fichier inconnu")
    '    Debug.Assert(bcolLignesLoaded, "Les lignes doivent être chargées")

    '    Dim nvcEntete As System.Collections.Specialized.NameValueCollection
    '    Dim bReturn As Boolean
    '    Dim nFile As Integer
    '    Dim objLg As LgFactTRP
    '    Dim n As Integer
    '    Dim nNombreChamps As Integer
    '    Dim strGuillemets As String
    '    Dim strValeur As String
    '    Dim strAttribut As String
    '    Try
    '        nFile = FreeFile()
    '        FileOpen(nFile, strFile, OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
    '        nvcEntete = ConfigurationManager.AppSettings()

    '        'Lecture de la Configuration de l'entete
    '        'Nombre de champs de l'entete
    '        nNombreChamps = nvcEntete("EXPORT_ENTETE_NOMBRE")
    '        For n = 1 To nNombreChamps
    '            'Guillemets Oui/non
    '            strGuillemets = nvcEntete("EXPORT_ENTETE_CHAMPS" & n & "_GUILLEMETS")
    '            'Valeur ou Attribut
    '            strValeur = nvcEntete("EXPORT_ENTETE_CHAMPS" & n & "_VALEUR")
    '            If strGuillemets.Equals("O") Then
    '                Print(nFile, Chr(34))
    '            End If
    '            If strValeur.Equals("ATTRIBUT") Then
    '                'Lecture du code de l'attribut
    '                strAttribut = nvcEntete("EXPORT_ENTETE_CHAMPS" & n & "_ATTRIBUT")
    '                'Decodage
    '                exporterAttribut(nFile, strAttribut, Me)
    '            Else
    '                Print(nFile, strValeur)
    '            End If

    '            If strGuillemets.Equals("O") Then
    '                Print(nFile, Chr(34))
    '            End If
    '            If n < nNombreChamps Then
    '                Print(nFile, ";")
    '            End If
    '        Next n
    '        PrintLine(nFile, "")

    '        'Lecture de la Configuration des lignes
    '        'Nombre de champs de l'entete
    '        nNombreChamps = nvcEntete("EXPORT_LIGNE_NOMBRE")

    '        For Each objLg In colLignes
    '            For n = 1 To nNombreChamps
    '                'Guillemets Oui/non
    '                strGuillemets = nvcEntete("EXPORT_LIGNE_CHAMPS" & n & "_GUILLEMETS")
    '                'Valeur ou Attribut
    '                strValeur = nvcEntete("EXPORT_LIGNE_CHAMPS" & n & "_VALEUR")
    '                If strGuillemets.Equals("O") Then
    '                    Print(nFile, Chr(34))
    '                End If
    '                If strValeur.Equals("ATTRIBUT") Then
    '                    'Lecture du code de l'attribut
    '                    strAttribut = nvcEntete("EXPORT_LIGNE_CHAMPS" & n & "_ATTRIBUT")
    '                    'Decodage
    '                    exporterAttribut(nFile, strAttribut, objLg)
    '                Else
    '                    Print(nFile, strValeur)
    '                End If

    '                If strGuillemets.Equals("O") Then
    '                    Print(nFile, Chr(34))
    '                End If
    '                If n < nNombreChamps Then
    '                    Print(nFile, ";")
    '                End If
    '            Next n
    '            PrintLine(nFile, "")
    '        Next objLg

    '        'Première ligne d'adresse
    '        nNombreChamps = nvcEntete("EXPORT_ADRLIV_NOMBRE")
    '        For n = 1 To nNombreChamps
    '            'Guillemets Oui/non
    '            strGuillemets = nvcEntete("EXPORT_ADRLIV_CHAMPS" & n & "_GUILLEMETS")
    '            'Valeur ou Attribut
    '            strValeur = nvcEntete("EXPORT_ADRLIV_CHAMPS" & n & "_VALEUR")
    '            If strGuillemets.Equals("O") Then
    '                Print(nFile, Chr(34))
    '            End If
    '            If strValeur.Equals("ATTRIBUT") Then
    '                'Lecture du code de l'attribut
    '                strAttribut = nvcEntete("EXPORT_ADRLIV_CHAMPS" & n & "_ATTRIBUT")
    '                'Decodage
    '                exporterAttribut(nFile, strAttribut, Me)
    '            Else
    '                Print(nFile, strValeur)
    '            End If

    '            If strGuillemets.Equals("O") Then
    '                Print(nFile, Chr(34))
    '            End If
    '            If n < nNombreChamps Then
    '                Print(nFile, ";")
    '            End If
    '        Next n
    '        PrintLine(nFile, "")

    '        'Deuxième ligne d'adresse
    '        nNombreChamps = nvcEntete("EXPORT_ADRFACT_NOMBRE")
    '        For n = 1 To nNombreChamps
    '            'Guillemets Oui/non
    '            strGuillemets = nvcEntete("EXPORT_ADRFACT_CHAMPS" & n & "_GUILLEMETS")
    '            'Valeur ou Attribut
    '            strValeur = nvcEntete("EXPORT_ADRFACT_CHAMPS" & n & "_VALEUR")
    '            If strGuillemets.Equals("O") Then
    '                Print(nFile, Chr(34))
    '            End If
    '            If strValeur.Equals("ATTRIBUT") Then
    '                'Lecture du code de l'attribut
    '                strAttribut = nvcEntete("EXPORT_ADRFACT_CHAMPS" & n & "_ATTRIBUT")
    '                'Decodage
    '                exporterAttribut(nFile, strAttribut, Me)
    '            Else
    '                Print(nFile, strValeur)
    '            End If

    '            If strGuillemets.Equals("O") Then
    '                Print(nFile, Chr(34))
    '            End If
    '            If n < nNombreChamps Then
    '                Print(nFile, ";")
    '            End If
    '        Next n
    '        PrintLine(nFile, "")

    '        FileClose(nFile)
    '        bReturn = True
    '    Catch ex As Exception
    '        bReturn = False
    '        setError("FactTRP.exporter", ex.ToString())
    '    End Try
    '    Return bReturn
    'End Function
    'Private Function exporterAttribut(ByVal nFile As Integer, ByVal strAttribut As String, ByVal obj As Object) As Boolean
    '    Dim objFact As FactTRP = Nothing
    '    Dim objLgFact As LgFactTRP = nothing

    '    Try
    '        objFact = CType(obj, FactTRP)
    '    Catch ex As Exception
    '        Try
    '            objLgFact = CType(obj, LgFactTRP)
    '        Catch ex1 As Exception
    '            Debug.Assert(False, "Objet de type FactTRP ou LGFactTRP requis")
    '            setError("FactTRP.exporterAttribut", "Objet de type FactTRP ou LGFactTRP requis")
    '            Return False
    '        End Try


    '    End Try
    '    Dim bReturn As Boolean
    '    Try
    '        bReturn = True
    '        Select Case strAttribut
    '            Case "FTRP_ID"
    '                Print(nFile, objFact.id)
    '            Case "FTRP_CODE"
    '                Print(nFile, objFact.code)
    '            Case "FTRP_DATE"
    '                Print(nFile, objFact.dateCommande.ToShortDateString)
    '            Case "FTRP_CLT_ID"
    '                Print(nFile, objFact.oTiers.id)
    '            Case "FTRP_TOTAL_TAXES"
    '                Print(nFile, objFact.montantTaxes)
    '            Case "FTRP_TOTAL_HT"
    '                Print(nFile, objFact.totalHT)
    '            Case "FTRP_TOTAL_TTC"
    '                Print(nFile, objFact.totalTTC)
    '            Case "FTRP_IDMODEREGLEMENT"
    '                Print(nFile, objFact.idModeReglement)
    '            Case "FTRP_PERIODE"
    '                Print(nFile, objFact.periode)
    '            Case "FTRP_ETAT"
    '                Print(nFile, objFact.etat.codeEtat)
    '            Case "FTRP_MONTANT_REGLEMENT"
    '                Print(nFile, objFact.montantReglement)
    '            Case "FTRP_DATE_REGLEMENT"
    '                Print(nFile, objFact.dateReglement)
    '            Case "FTRP_REF_REGLEMENT"
    '                Print(nFile, objFact.refReglement)
    '            Case "FTRP_COM_FACT"
    '                Print(nFile, objFact.CommFacturation)
    '            Case "FTRP_DATE_REGLEMENT"
    '                Print(nFile, objFact.dateReglement)
    '            Case "CLT_CODE"
    '                Print(nFile, objFact.oTiers.code)
    '            Case "CLT_RS"
    '                Print(nFile, objFact.oTiers.rs)
    '            Case "CLT_ADRLIV_RUE1"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.rue1)
    '            Case "CLT_ADRLIV_RUE2"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.rue2)
    '            Case "CLT_ADRLIV_RUE1+2"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.rue1 & vbCrLf & objFact.oTiers.AdresseLivraison.rue2)
    '            Case "CLT_ADRLIV_CP"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.cp)
    '            Case "CLT_ADRLIV_VILLE"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.ville)
    '            Case "CLT_ADRLIV_TEL"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.tel)
    '            Case "CLT_ADRLIV_FAX"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.fax)
    '            Case "CLT_ADRLIV_PORT"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.port)
    '            Case "CLT_ADRLIV_EMAIL"
    '                Print(nFile, objFact.oTiers.AdresseLivraison.Email)
    '            Case "CLT_ADRFACT_RUE1"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.rue1)
    '            Case "CLT_ADRFACT_RUE2"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.rue2)
    '            Case "CLT_ADRFACT_RUE1+2"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.rue1 & vbCrLf & objFact.oTiers.AdresseFacturation.rue2)
    '            Case "CLT_ADRFACT_CP"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.cp)
    '            Case "CLT_ADRFACT_VILLE"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.ville)
    '            Case "CLT_ADRFACT_TEL"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.tel)
    '            Case "CLT_ADRFACT_FAX"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.fax)
    '            Case "CLT_ADRFACT_PORT"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.port)
    '            Case "CLT_ADRFACT_EMAIL"
    '                Print(nFile, objFact.oTiers.AdresseFacturation.Email)
    '            Case "LIB_MODEREGLEMENT"
    '                Dim oParam As Param
    '                oParam = New Param
    '                oParam.load(m_idModeRegelement)
    '                Print(nFile, oParam.valeur)
    '            Case "ID_TAUXTVA"
    '                Print(nFile, idParamTVA)
    '            Case "LIB_TAUXTVA"
    '                Dim oParam As Param
    '                oParam = New Param
    '                oParam.load(m_idParamTVA)
    '                Print(nFile, oParam.valeur)

    '            Case "LGTRP_ID"
    '                Print(nFile, objLgFact.id)
    '            Case "LGTRP_NUM"
    '                Print(nFile, objLgFact.num)
    '            Case "LGTRP_CODEMAXIMA"
    '                'Attention le code MAXIMA est calculé
    '                If (objLgFact.num = CST_LGFACTTRP_NUM_GO) Then
    '                    Print(nFile, ConfigurationManager.AppSettings.GetValues("EXPORT_CODE_GAZOLE")(0))
    '                Else
    '                    Print(nFile, ConfigurationManager.AppSettings.GetValues("EXPORT_CODE_TRANSPORT")(0))
    '                End If
    '            Case "LGTRP_FACTTRP_ID"
    '                Print(nFile, objLgFact.idFactTRP)
    '            Case "LGTRP_CMDCLT_ID"
    '                Print(nFile, objLgFact.idCmdCLT)
    '            Case "LGTRP_TRPEUR"
    '                Print(nFile, objLgFact.nomTransporteur)
    '            Case "LGTRP_DATE_LIV"
    '                Print(nFile, objLgFact.dateLivraison)
    '            Case "LGTRP_REF_LIV"
    '                Print(nFile, objLgFact.referenceLivraison)
    '            Case "LGTRP_DATE_CMD"
    '                Print(nFile, objLgFact.dateCommande)
    '            Case "LGTRP_REF_CMD"
    '                Print(nFile, objLgFact.refCommande)
    '            Case "LGTRP_QTE_COLIS"
    '                Print(nFile, objLgFact.qteColis)
    '            Case "LGTRP_QTE_POIDS"
    '                Print(nFile, objLgFact.poids)
    '            Case "LGTRP_QTE_PAL_NONPREP"
    '                Print(nFile, objLgFact.qtePalettesNonPreparees)
    '            Case "LGTRP_QTE_PAL_PREP"
    '                Print(nFile, objLgFact.qtePalettesPreparees)
    '            Case "LGTRP_PU_PAL_NONPREP"
    '                Print(nFile, objLgFact.puPalettesNonPreparees)
    '            Case "LGTRP_PU_PAL_PREP"
    '                Print(nFile, objLgFact.puPalettesPreparees)
    '            Case "LGTRP_MT_HT"
    '                Print(nFile, objLgFact.prixHT)
    '            Case "LGTRP_RESUME"
    '                Print(nFile, objLgFact.dateLivraison.ToShortDateString & " " & objLgFact.referenceLivraison & " " & objLgFact.nomTransporteur & " " & objLgFact.dateCommande.ToShortDateString & " " & objLgFact.refCommande)
    '            Case Else
    '                Print(nFile, strAttribut)
    '        End Select
    '        bReturn = True
    '    Catch ex As Exception
    '        bReturn = False
    '        setError("FactTRP.exporterAttribut", ex.ToString)
    '    End Try
    '    Debug.Assert(bReturn, getErreur())

    '    Return bReturn
    'End Function

    Public Overrides Sub Exporter(ByVal pstrFileName As String)

        MyBase.ExporterFacture(pstrFileName, "TRP")
    End Sub

End Class
