Imports System.configuration
Public Class FactColisageJ
    Inherits Facture

    Private m_periode As String
    Private m_montantTaxes As Decimal
    Private m_montantReglement As Decimal
    Private m_refReglement As String
    Private m_dateReglement As Date
    Private m_dossier As String
    '    Private m_idModeRegelement As Long

#Region "Accesseurs"
    Public Sub New(ByVal poFournisseur As Fournisseur)
        MyBase.New(poFournisseur, EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactCOLGeneree))
        m_periode = Now.ToString("MMMM yyyy")
        m_montantTaxes = Param.getConstante("CST_FACT_COL_TAXES")
        m_montantReglement = 0
        m_refReglement = ""
        m_dateReglement = DATE_DEFAUT
        m_typedonnee = vncEnums.vncTypeDonnee.FACTCOL
        '       m_idModeRegelement = Param.getConstante("CST_COL_IDMODEREGLEMENT")
        m_TotalHT = 0
        m_TotalTTC = 0
        m_IDModeReglement = poFournisseur.idModeReglement2 'utilisation du mode de reglement du Tiers
        m_dossier = Dossier.VINICOM

    End Sub
    Public Shared Function createandload(ByVal pid As Long) As FactColisageJ
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactColisageJ
        obj = New FactColisageJ
        Try
            If pid <> 0 Then
                obj.load(pid)
            End If
        Catch ex As Exception
            setError("FactColisage.CreateandLoad", ex.ToString)
        End Try
        Return obj
    End Function 'createandload
    Public Shared Function createandload(ByVal pCode As String) As FactColisageJ
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactColisageJ = Nothing
        Dim oCol As Collection
        Try
            obj = New FactColisageJ()
            oCol = getListe(pCode)
            If (oCol.Count = 1) Then
                obj = oCol(1)
            End If

        Catch ex As Exception
            setError("CreateFactCol", ex.Message)
        End Try
        Debug.Assert(obj.code = pCode, "FactColisage " & pCode & " non chargée")
        Return obj
    End Function 'createandload
    Friend Sub New()
        Me.New(New Fournisseur())
        m_dossier = Dossier.VINICOM
    End Sub
    'Public Property idModeReglement() As Long
    '    Get
    '        Debug.Assert(Not bResume, "Objet Résumé")
    '        Return m_idModeRegelement
    '    End Get
    '    Set(ByVal Value As Long)
    '        If Value <> m_idModeRegelement Then
    '            RaiseUpdated()
    '            m_idModeRegelement = Value
    '        End If
    '    End Set
    'End Property
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

    Public Property periode() As String
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_periode
        End Get
        Set(ByVal Value As String)
            If Value <> m_periode Then
                Try
                    RaiseUpdated()
                    m_periode = CDate(Value).ToString("MMMM yyyy")
                Catch ex As Exception

                End Try
            End If
        End Set

    End Property
    Public Property dossierFact() As String
        Get
            Debug.Assert(Not bResume, "Objet Résumé")
            Return m_dossier
        End Get
        Set(ByVal Value As String)
            If Value <> m_dossier Then
                RaiseUpdated()
                m_dossier = Value
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
    Public Property Commentaire() As String
        Get
            Return CommFacturation.comment
        End Get
        Set(ByVal Value As String)
            If Value <> CommFacturation.comment Then
                RaiseUpdated()
                CommFacturation.comment = Value
            End If
        End Set

    End Property


    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return code & " | " & oTiers.rs & " | " & dateFacture.ToShortDateString & " | " & Format(totalTTC, "C") & " | " & etat.libelle
        End Get
    End Property



    Public ReadOnly Property dateDebut As Date
        Get
            Try
                Return CDate(periode)
            Catch ex As Exception
                Return CDate(Now.ToString("MMMM yyyy"))
            End Try
        End Get
    End Property

    Public ReadOnly Property dateFin As Date
        Get
            Dim dDeb As Date = dateDebut
            Return dDeb.AddMonths(1).AddDays(-1)
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
        Dim objCommande As FactColisageJ

        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj
            bReturn = bReturn And (m_periode.Equals(objCommande.periode))

            bReturn = bReturn And (m_montantReglement.Equals(objCommande.montantReglement))
            bReturn = bReturn And (m_dateReglement.ToShortDateString().Equals(objCommande.dateReglement.ToShortDateString()))
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
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        If bDeleted Then

            bReturn = supprimeMvtStock()
        End If
        bReturn = MyBase.Save()
        If m_id > 0 Then
            'MVTS STOCKS
            If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncGenerer Then
                bReturn = bReturn And genereMvtStock()
                Debug.Assert(bReturn, "Erreur en generemvtStock" & getErreur())
            End If
        End If
        shared_disconnect()
        Return bReturn
    End Function ' Save
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean = False
        Try
            If pid <> 0 Then
                m_id = pid
            End If

            Debug.Assert(id <> 0, "idFactCOL <> 0")
            bReturn = loadFactColisage()
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

        Debug.Assert(bReturn, "FactCol.DBLoad: " & getErreur())
        Return bReturn
    End Function 'DBLoad

    Friend Overrides Function delete() As Boolean

        Dim bReturn As Boolean = False
        Try
            bReturn = True
            If bcolLignesLoaded Then
                bReturn = deletecolLgCOL()
            End If
            If bReturn Then
                bReturn = deleteFactColisage()
            End If
            If bReturn Then
                ' Libération des Mvts de stocks
                bReturn = supprimeMvtStock()
            End If
            deleteReglements()
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactColisage.delete: " & getErreur())
        Return bReturn

    End Function 'delete

    ''' <summary>
    ''' Affection de la référence de facture de colisage aux mvt de stock
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function genereMvtStock() As Boolean

        Dim olgFactCol As LgFactColisage
        Dim oColmvt As New Collection()
        Dim oFournisseur As Fournisseur
        Dim omvtStock As mvtStock
        Dim objPRD As Produit

        oFournisseur = CType(oTiers, Fournisseur)

        If Not m_bcolLgLoaded Then
            loadcolLignes()
        End If



        'Pour chaque liggne d'une facture de colisage (Normalement 1 seule
        For Each olgFactCol In colLignes
            'Recupération de la Liste des mouvements de stocks
            If dossierFact = Dossier.VINICOM Then
                oColmvt = mvtStock.getListe2(dateDebut, dateFin, oFournisseur, vncEtatMVTSTK.vncMVTSTK_nFact)
            End If
            If dossierFact = Dossier.HOBIVIN Then
                oColmvt = mvtStock.getListeDossierNonFacture(dossierFact, dateDebut, dateFin)
            End If

            For Each omvtStock In oColmvt
                objPRD = New Produit()
                objPRD.DBLoadLight(omvtStock.idProduit)
                If (objPRD.bStock) Then
                    omvtStock.idFactColisage = Me.id
                    omvtStock.changeEtat(vncActionFactColisage.vncActionFacturer)
                Else
                    omvtStock.idFactColisage = 0

                End If
                omvtStock.save()
            Next
        Next

        Return True

    End Function

    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Producteur n'est pas Renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Producteur n'est pas sauvegardé")
        Debug.Assert(id = 0, "id= 0")

        Dim bReturn As Boolean = False

        Try
            bReturn = False
            If setNewcode() Then
                bReturn = insertFactColisage()
            End If
            If bcolLignesLoaded Then
                bReturn = savecolLignes()
            End If
        Catch ex As Exception
            bReturn = False
            setError("FactColisage.insert", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactColisage.insert: " & getErreur())
        Return bReturn

    End Function 'insert

    ''' <summary>
    ''' Libère les mouvements de stocks de la facture de colisade
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function supprimeMvtStock() As Boolean
        Dim olgFactCol As LgFactColisage
        Dim oColmvt As New Collection()
        Dim oFournisseur As Fournisseur
        Dim omvtStock As mvtStock

        oFournisseur = CType(oTiers, Fournisseur)

        If Not m_bcolLgLoaded Then
            loadcolLignes()
        End If

        'Pour chaque liggne d'une facture de colisage (Normalement 1 seule
        For Each olgFactCol In colLignes
            'Recupération de la Liste des mouvements de stocks
            If dossierFact = Dossier.VINICOM Then
                oColmvt = mvtStock.getListe2(dateDebut, dateFin, oFournisseur, vncEtatMVTSTK.vncMVTSTK_Fact)
            End If
            If dossierFact = Dossier.HOBIVIN Then
                oColmvt = mvtStock.getListeDossierFacture(dossierFact, dateDebut, dateFin)
            End If
            For Each omvtStock In oColmvt
                omvtStock.idFactColisage = 0
                omvtStock.changeEtat(vncActionFactColisage.vncActionAnnFacturer)
                omvtStock.save()
            Next
        Next
        Return True
    End Function

    Friend Overrides Function update() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Producteur n'est pas renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Producteur n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idFact <> 0")
        Dim bReturn As Boolean = False
        Try
            'Mise à jour de la sous-commande
            bReturn = UpdateFactColisage()
            If bcolLignesLoaded Then
                bReturn = savecolLignes()
            End If
        Catch ex As Exception
            bReturn = False
            setError("FactColisage.Update", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactColisage.Update: " & getErreur())
        Return bReturn

    End Function

    Public Overrides Function loadcolLignes() As Boolean
        Debug.Assert(m_id <> 0, "La Facture doit être sauvegardé au Préalable")
        Dim bReturn As Boolean = False
        shared_connect()
        If m_bcolLgLoaded Then
            m_colLignes.clear()
        End If
        bReturn = LoadcolLgFactColisage()

        Debug.Assert(bReturn, ColEvent.getErreur())
        m_bcolLgLoaded = bReturn ' Les lignes sont chargées
        m_bColLgInsertorDelete = False
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
        Debug.Assert(m_id <> 0, "La Facture  doit être sauvegardée au préalable")
        'En mode commandeClient ilfaut supprimer les lignes avant de les recréer pour gérer les suppressions et ajouts de lignes.
        If m_bColLgInsertorDelete Then
            'On Supprime la collection des lignes avant de la recréer 
            '1° Suppression de la reference à la facture de transport dans la table Commande
            bReturn = deletecolLgCOL()
            If bReturn Then
                bReturn = INSERTcolLGCOL()
            End If
            m_bColLgInsertorDelete = False
        Else
            'On Met à jour la collection (SousCommande)
            bReturn = UPDATEcolLGCOL()
        End If
        Debug.Assert(bReturn, "FactColisage.savecolLignes:" & getErreur())
        Return bReturn
    End Function

    Public Function purge() As Boolean
        '================================================================================
        ' Fonction : purge 
        ' Description : Supression de la facture ET de tous ses élements dépendants
        '================================================================================
        Dim bReturn As Boolean = False
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
        ncode = GetNumeroFactureColisage()
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
        Dim oLg As LgFactColisage
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
    Public Overrides Function calculPrixTotal() As Boolean
        Dim oLg As LgFactColisage = Nothing
        Dim oLgGO As LgFactColisage = Nothing
        Dim nHT As Decimal
        Dim nHTCmd As Decimal
        Dim nTTC As Decimal
        Dim objParam As Param

        nHT = 0
        nHTCmd = 0
        nTTC = 0
        For Each oLg In m_colLignes
            nHT = nHT + oLg.prixHT
        Next oLg


        If Param.getConstante("CST_FACT_COL_TAXES") <> 0 Then
            montantTaxes = Param.getConstante("CST_FACT_COL_TAXES")
        End If

        nHT = nHT + montantTaxes

        objParam = New Param
        objParam.load(m_idParamTVA)

        nTTC = Math.Round(nHT * (1 + (objParam.valeur / 100)), 2)
        'On passe par les accesseurs pour lever l'event Updated s'il y a lieu
        totalHT = nHT
        totalTTC = nTTC
        Return True
    End Function 'calculPrixTotal

    Public Overrides Function changeEtat(ByVal p_action As vncActionEtatCommande) As Boolean
        Debug.Assert(p_action >= vncActionEtatCommande.vncActionMinFactCOL And p_action <= vncActionEtatCommande.vncActionMaxFactCOL)
        Dim bReturn As Boolean = False
        Try
            If p_action >= vncActionEtatCommande.vncActionMinFactCOL And p_action <= vncActionEtatCommande.vncActionMaxFactCOL Then
                MyBase.changeEtat(p_action)
                RaiseUpdated()
                bReturn = True
            Else
                setError("FactCol.changeEtat", "Code Action incorrect :" + p_action)
                bReturn = False
            End If

        Catch ex As Exception
            setError("FactCol.changeEtat", ex.ToString)
            bReturn = False
        End Try

        Return bReturn
    End Function 'ChangeEtat

    Public Overrides Function ControleMvtStock() As Microsoft.VisualBasic.Collection
        Return Nothing
    End Function

    Public Overrides Sub exporter(pstrFileName As String)
        MyBase.ExporterFacture(pstrFileName, "COL")
    End Sub

    Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not p_oProduit Is Nothing)

        Dim oLgFact As LgFactColisage

        oLgFact = New LgFactColisage(m_id)
        oLgFact.num = p_strNum
        oLgFact.idCmd = id
        oLgFact.idSCmd = 0 ' La sous commande n'est pas connue
        oLgFact.oProduit = p_oProduit
        oLgFact.qteCommande = p_qteCmd
        oLgFact.prixU = p_prixU
        oLgFact.bGratuit = p_bGratuit
        oLgFact.prixHT = p_prixHT
        oLgFact.prixTTC = p_prixTTC

        oLgFact = AjouteLigne(oLgFact, p_bCalculPrix)
        Return oLgFact
    End Function 'AjouteLigne

#End Region
    Public Shared Function getListe(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pCodeFournisseur As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection = New Collection()

        shared_connect()
        colReturn = Persist.ListeFACTColisage(pddeb, pdfin, pCodeFournisseur, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCol.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function

    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNomClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection = New Collection()

        shared_connect()
        colReturn = Persist.ListeFACTColisage(strCode, strNomClient, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCol.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function
    Public Shared Function getListeNonReglee(Optional ByVal strCode As String = "", Optional ByVal strNomFournisseur As String = "") As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeFACTColisageNonReglee(strCode, strNomFournisseur)
        Debug.Assert(Not colReturn Is Nothing, "FactColisage.getListeNonReglee" & getErreur())
        shared_disconnect()
        Return colReturn
    End Function

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
        Dim objLg As LgFactColisage

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
                setError("FactCol.SupprimeLigne(" & strNumLigne & ")", "Ligne inconnue")
                bReturn = False
            End If
        Catch ex As Exception
            setError("FactCol.SupprimeLigne(" & strNumLigne & ")", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SupprimeLigneCol

    Public Shared Function GenereFacture(ByVal pDate As Date, ByVal pFourn As Fournisseur, pDossier As String) As FactColisageJ

        Dim oFactCol As FactColisageJ = Nothing
        Dim oLgFactCol As LgFactColisage = Nothing
        Dim oDS As dsVinicom
        Dim oRow As dsVinicom.RECAPCOLISAGEJOURNRow
        Dim dDatePrec As Date = DateTime.MinValue
        Dim pfin As Date
        Dim pDeb As Date
        pDeb = New Date(pDate.Year, pDate.Month, 1)
        pfin = pDeb.AddMonths(1).AddDays(-1)

        ' Recupère la liste des mouvements de stocks non facturé
        oDS = GenereDataSetRecapColisage(pDeb, pfin, pFourn.code, Param.getConstante("CST_FACT_COL_PU_COLIS"), pDossier)
        'Dim oRECAPRow As dsVinicom.RECAP_COLISAGERow

        'For Each oRECAPRow In oDS.RECAP_COLISAGE
        '    Console.WriteLine(oRECAPRow.RC_DATE.ToShortDateString() + "," + oRECAPRow.RC_TYPE + "," + oRECAPRow.RC_SI.ToString() + "+" + oRECAPRow.RC_ENTREE.ToString() + "-" + oRECAPRow.RC_SORTIE.ToString() + " = " + oRECAPRow.RC_SF.ToString() + "/" + oRECAPRow.RC_COUT.ToString())
        'Next

        oFactCol = New FactColisageJ(pFourn)
        oFactCol.dossierFact = pDossier
        oFactCol.periode = pDeb.ToString("MMMM yyyy")

        For Each oRow In oDS.RECAPCOLISAGEJOURN

            Dim qte As Decimal
            qte = oRow.RC_S01 + oRow.RC_S02 + oRow.RC_S03 + oRow.RC_S04 + oRow.RC_S05 + oRow.RC_S06 + oRow.RC_S07 + oRow.RC_S08 + oRow.RC_S09 + oRow.RC_S10 + oRow.RC_S11 + oRow.RC_S12 + oRow.RC_S13 + oRow.RC_S14 + oRow.RC_S15 + oRow.RC_S16 + oRow.RC_S17 + oRow.RC_S18 + oRow.RC_S19 + oRow.RC_S20 + oRow.RC_S21 + oRow.RC_S22 + oRow.RC_S23 + oRow.RC_S24 + oRow.RC_S25 + oRow.RC_S26 + oRow.RC_S27 + oRow.RC_S28 + oRow.RC_S29 + oRow.RC_S30 + oRow.RC_S31
            Dim oProduit As Produit
            oProduit = New Produit()
            oProduit.load(oRow.RC_IDPRODUIT)
            If oProduit.id <> 0 Then
                If qte <> 0 Then
                    Dim nNum As Integer = oFactCol.getNextNumLg()
                    oFactCol.AjouteLigne(nNum, p_oProduit:=oProduit, p_qteCmd:=qte, p_prixU:=oRow.RC_COUT_U)
                End If
            End If
        Next
        Return oFactCol

    End Function 'GenereFacture

    Public Shared Function GenereDataSetRecapColisage(ByVal pdDeb As Date, ByVal pdFin As Date, ByVal pCodeFourn As String, ByVal pCout As Decimal, pdossier As String) As dsVinicom



        Dim dDatePrec As Date = DateTime.MinValue
        Dim pDS As dsVinicom
        Dim colPRD As New Collection()


        pDS = New dsVinicom()


        Dim idFourn As Integer = 0
        If Not String.IsNullOrEmpty(pCodeFourn) And pdossier = Dossier.VINICOM Then
            Dim ofrn As Fournisseur
            ofrn = Fournisseur.createandload(pCodeFourn)
            idFourn = ofrn.id
        End If


        Try
            If pdossier = Dossier.VINICOM Then
                'Charegement de la Liste des produits Plateformes
                colPRD = Produit.getListe(vncTypeProduit.vncPlateforme, idFournisseur:=idFourn)
            End If
            If pdossier = Dossier.HOBIVIN Then
                'Charegement de la Liste des produits Plateformes
                colPRD = Produit.getListe(vncTypeProduit.vncPlateforme, pdossier:=pdossier)
            End If
            'chargement des Mouvements de Stocks depuis le dernier inventaire et Calcul du Stock initial
            For Each oPRD As Produit In colPRD
                If oPRD.bDisponible Then
                    oPRD.load()
                    oPRD.loadcolmvtStockDepuisLeDernierMouvementInventaire()
                    oPRD.GenereDataSetRecapColisage(pdDeb, pdFin, pCout, pDS)
                End If
            Next

            Dim oRow As dsVinicom.CONSTANTESRow
            oRow = pDS.CONSTANTES.NewCONSTANTESRow()
            oRow.CST_SOC2_ADRESSE_RUE1 = Param.getConstante("CST_SOC2_ADRESSE_RUE1")
            oRow.CST_SOC2_ADRESSE_RUE2 = Param.getConstante("CST_SOC2_ADRESSE_RUE2")
            oRow.CST_SOC2_ADRESSE_CP = Param.getConstante("CST_SOC2_ADRESSE_CP")
            oRow.CST_SOC2_ADRESSE_VILLE = Param.getConstante("CST_SOC2_ADRESSE_VILLE")
            oRow.CST_SOC2_TEL = Param.getConstante("CST_SOC2_TEL")
            oRow.CST_SOC2_FAX = Param.getConstante("CST_SOC2_FAX")
            oRow.CST_SOC2_EMAIL = Param.getConstante("CST_SOC2_EMAIL")

            pDS.CONSTANTES.AddCONSTANTESRow(oRow)
        Catch ex As Exception
            setError(ex.StackTrace, ex.Message)
            pDS = New dsVinicom()
        End Try


        Return pDS

    End Function
    Public Function GenereDataSetRecapColisage() As dsVinicom

        Dim dDatePrec As Date = DateTime.MinValue
        Dim pDS As dsVinicom
        Dim colPRD As New Collection()


        pDS = New dsVinicom()


        Try
            For Each oLg As LgFactColisage In colLignes
                Dim oPrd As Produit
                oPrd = oLg.oProduit
                oPrd.load()

                oPrd.loadcolmvtStockDepuisLeDernierMouvementInventaire()
                oPrd.GenereDataSetRecapColisage(dateDebut, dateFin, oLg.prixU, pDS)
            Next

            Dim oRow As dsVinicom.CONSTANTESRow
            oRow = pDS.CONSTANTES.NewCONSTANTESRow()
            oRow.CST_SOC2_ADRESSE_RUE1 = Param.getConstante("CST_SOC2_ADRESSE_RUE1")
            oRow.CST_SOC2_ADRESSE_RUE2 = Param.getConstante("CST_SOC2_ADRESSE_RUE2")
            oRow.CST_SOC2_ADRESSE_CP = Param.getConstante("CST_SOC2_ADRESSE_CP")
            oRow.CST_SOC2_ADRESSE_VILLE = Param.getConstante("CST_SOC2_ADRESSE_VILLE")
            oRow.CST_SOC2_TEL = Param.getConstante("CST_SOC2_TEL")
            oRow.CST_SOC2_FAX = Param.getConstante("CST_SOC2_FAX")
            oRow.CST_SOC2_EMAIL = Param.getConstante("CST_SOC2_EMAIL")

            pDS.CONSTANTES.AddCONSTANTESRow(oRow)
        Catch ex As Exception
            setError(ex.StackTrace, ex.Message)
            pDS = New dsVinicom()
        End Try


        Return pDS

    End Function

End Class
