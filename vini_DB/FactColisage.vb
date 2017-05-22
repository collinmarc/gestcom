Imports System.configuration
Public Class FactColisage
    Inherits Facture

    Private m_periode As String
    Private m_montantTaxes As Decimal
    Private m_montantReglement As Decimal
    Private m_refReglement As String
    Private m_dateReglement As Date
    '    Private m_idModeRegelement As Long

#Region "Accesseurs"
    Public Sub New(ByVal poFournisseur As Fournisseur)
        MyBase.New(poFournisseur, EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactCOLGeneree))
        m_periode = ""
        m_montantTaxes = Param.getConstante("CST_FACT_COL_TAXES")
        m_montantReglement = 0
        m_refReglement = ""
        m_dateReglement = DATE_DEFAUT
        m_typedonnee = vncEnums.vncTypeDonnee.FACTCOL
        '       m_idModeRegelement = Param.getConstante("CST_COL_IDMODEREGLEMENT")
        m_TotalHT = 0
        m_TotalTTC = 0
        idModeReglement = poFournisseur.idModeReglement2 'utilisation du mode de reglement du Tiers
    End Sub
    Public Shared Function createandload(ByVal pid As Long) As FactColisage
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactColisage
        obj = New FactColisage
        Try
            If pid <> 0 Then
                obj.load(pid)
            End If
        Catch ex As Exception
            setError("FactColisage.CreateandLoad", ex.ToString)
        End Try
        Return obj
    End Function 'createandload
    Public Shared Function createandload(ByVal pCode As String) As FactColisage
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactColisage = Nothing
        Dim oCol As Collection
        Try
            obj = New FactColisage()
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
        Dim objCommande As FactColisage

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
        bReturn = MyBase.Save()
        If m_id > 0 Then
            'MVTS STOCKS
            If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncGenerer Then
                bReturn = bReturn And genereMvtStock()
                Debug.Assert(bReturn, "Erreur en generemvtStock" & getErreur())
            End If
            If getActionMvtStock() = vncEnums.vncGenererSupprimer.vncSupprimer Then
                bReturn = bReturn And supprimeMvtStock()
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
        Dim oColmvt As Collection
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
            oColmvt = mvtStock.getListe2(olgFactCol.dDeb, olgFactCol.dFin, oFournisseur, vncEtatMVTSTK.vncMVTSTK_nFact)


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
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
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
        Dim oColmvt As Collection
        Dim oFournisseur As Fournisseur
        Dim omvtStock As mvtStock

        oFournisseur = CType(oTiers, Fournisseur)

        If Not m_bcolLgLoaded Then
            loadcolLignes()
        End If

        'Pour chaque liggne d'une facture de colisage (Normalement 1 seule
        For Each olgFactCol In colLignes
            'Recupération de la Liste des mouvements de stocks
            oColmvt = mvtStock.getListe2(olgFactCol.dDeb, olgFactCol.dFin, oFournisseur, vncEtatMVTSTK.vncMVTSTK_Fact)
            For Each omvtStock In oColmvt
                omvtStock.idFactColisage = 0
                omvtStock.changeEtat(vncActionFactColisage.vncActionAnnFacturer)
                omvtStock.save()
            Next
        Next
        Return True
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
        Debug.Assert(m_id <> 0, "La commande  doit être sauvegardée au préalable")
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
    Public Overloads Function AjouteLigne(ByVal pobjLgCMD As LgCommande, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "Méthode désactivée dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "Méthode désactivée dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
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
            nHT = nHT + oLg.MontantHT
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

    Public Function AjouteLigneFactColisage(ByVal p_oLg As LgFactColisage, Optional ByVal pCalcul As Boolean = True) As LgFactColisage
        ''=======================================================================
        ''Fonction : AjouteLigneFactColisage()
        ''Description : Ajoute une ligne à la facture (permet d'ajouter une ligne qui n'a pas de commande)
        ''Retour : une ligne de commande ou nothing si l'ajout échoue
        ''=======================================================================
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not p_oLg Is Nothing)
        'Debug.Assert(Me.id <> 0, "La Facture doit être enregistrée")

        Dim oLg As LgFactColisage

        Try
            p_oLg.idFactColisage = Me.id
            colLignes.Add(p_oLg, p_oLg.num)
            If pCalcul Then
                calculPrixTotal()
            End If
            m_bcolLgLoaded = True
            m_bColLgInsertorDelete = True
            oLg = p_oLg
        Catch ex As Exception
            setError("AjouteLigneFactColisage", ex.ToString)
            oLg = Nothing
        End Try
        Debug.Assert(Not oLg Is Nothing, getErreur())
        Return oLg
    End Function 'AjouteLigneFactColisage
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

    Public Shared Function GenereFacture(ByVal pdInventaire As Date, ByVal pdfin As Date, ByVal pFourn As Fournisseur) As FactColisage

        Dim oFactCol As FactColisage = Nothing
        Dim oLgFactCol As LgFactColisage = Nothing
        Dim oDS As dsVinicom
        Dim oRow As dsVinicom.RECAP_COLISAGERow
        Dim dDatePrec As Date = DateTime.MinValue
        ' Recupère la liste des mouvements de stocks non facturé
        oDS = GenereDataSetRecapColisage(pdInventaire, pdfin, pFourn.code, Param.getConstante("CST_FACT_COL_PU_COLIS"))
        'Dim oRECAPRow As dsVinicom.RECAP_COLISAGERow

        'For Each oRECAPRow In oDS.RECAP_COLISAGE
        '    Console.WriteLine(oRECAPRow.RC_DATE.ToShortDateString() + "," + oRECAPRow.RC_TYPE + "," + oRECAPRow.RC_SI.ToString() + "+" + oRECAPRow.RC_ENTREE.ToString() + "-" + oRECAPRow.RC_SORTIE.ToString() + " = " + oRECAPRow.RC_SF.ToString() + "/" + oRECAPRow.RC_COUT.ToString())
        'Next

        oFactCol = New FactColisage(pFourn)


        For Each oRow In oDS.RECAP_COLISAGE

            If oRow.RC_TYPE = "9" Then
                If oFactCol Is Nothing Then
                    oFactCol = New FactColisage(pFourn)
                End If
                oFactCol.AjouteLigneFactColisage(getDateDebutdeMois(oRow.RC_DATE), getDateFindeMois(oRow.RC_DATE), oRow.RC_SI, oRow.RC_ENTREE, oRow.RC_SORTIE, oRow.RC_SF)
            End If
        Next
        Return oFactCol

    End Function 'GenereFacture
    ''' <summary>
    ''' Genere les lignes de facture pour les mois sans mouvements
    ''' </summary>
    ''' <param name="pFactCol"></param>
    ''' <param name="pDate"></param>
    ''' <param name="pDateLimite"></param>
    ''' <param name="pSI"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function genereLignesFactureMoisSansMouvements(ByVal pFactCol As FactColisage, ByVal pDate As Date, ByVal pDateLimite As Date, ByVal pSI As Integer) As Boolean
        Dim dDeb As Date
        Dim dFin As Date
        Dim bReturn As Boolean
        Try

            While (pDate.AddMonths(1) < pDateLimite)
                pDate = pDate.AddMonths(1)
                dDeb = New Date(pDate.Year, pDate.Month, 1)
                dFin = New Date(pDate.Year, pDate.Month, DateTime.DaysInMonth(pDate.Year, pDate.Month))
                pFactCol.AjouteLigneFactColisage(dDeb, dFin, pSI, 0, 0, pSI)
            End While
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn

    End Function

    Public Shared Function GenereDataSetRecapColisage(ByVal pdDeb As Date, ByVal pdFin As Date, ByVal pCodeFourn As String, ByVal pCout As Decimal) As dsVinicom

        'Debug.Assert(Not String.IsNullOrEmpty(pCodeFourn), "Le Code Fournisseur doit être spécifié")

        Dim oMvt As vini_DB.dsVinicom.MVT_STOCKRow
        Dim nSI As Decimal
        Dim nSF As Decimal
        Dim nEntree As Decimal
        Dim nSortie As Decimal
        Dim nEntreeMois As Decimal
        Dim nSortieMois As Decimal
        Dim oRow As dsVinicom.RECAP_COLISAGERow
        Dim oTAMvt As New vini_DB.dsVinicomTableAdapters.MVT_STOCKTableAdapter()

        Dim dDatePrec As Date = DateTime.MinValue
        Dim objFRN As Fournisseur
        Dim nQteColis As Decimal
        Dim pDS As dsVinicom
        Dim colFourn As Collection
        Dim bTrtCommence As Boolean '' Le traitement est -il commencé ?

        pDS = New dsVinicom()
        nSI = 0
        nSF = 0
        nEntree = 0
        nSortie = 0
        Try


            colFourn = Fournisseur.getListe(pCodeFourn)
            For Each objFRN In colFourn
                If objFRN.code <> "999" Then
                    oTAMvt.Connection = Persist.oleDBConnection
                    ' Liste des Mouvements de stock trié par FRN/DATE/TYPE/PRD_CODE
                    oTAMvt.FillForRECAPCOLISAGE(pDS.MVT_STOCK, pdDeb.ToShortDateString(), pdFin.ToShortDateString(), objFRN.code)
                    dDatePrec = DateTime.MinValue
                    nSI = 0
                    nSF = 0
                    nEntree = 0
                    nSortie = 0
                    nEntreeMois = 0
                    nSortieMois = 0
                    bTrtCommence = False

                    For Each oMvt In pDS.MVT_STOCK
                        'Les Mouvements d'inventaire autre que Ceux du debut ne sont pas pris en compte
                        If oMvt.STK_TYPE = vncTypeMvt.vncMvtInventaire And Not oMvt.STK_DATE.ToShortDateString.Equals(pdDeb.ToShortDateString) Then
                            Continue For
                        End If

                        'Calcul de la Quantité de Colis (Arrondi au Supéieur)
                        If oMvt.IsCONDNull() Then
                            nQteColis = 1
                        Else
                            nQteColis = oMvt.STK_QTE / oMvt.COND
                            If (nQteColis - Int(nQteColis) <> 0) Then
                                nQteColis = nQteColis + 1
                            End If
                        End If
                        oMvt.STK_QTECOLIS = Int(nQteColis)

                        'Test sur la rupture de mois pour calculer le cout
                        If bTrtCommence Then
                            If IsRuptureMois(oMvt.STK_DATE, dDatePrec) Then
                                AjoutLigneRecapMoisEnCours(pDS, objFRN, pCout, nEntreeMois, nSortieMois)
                                nEntreeMois = 0
                                nSortieMois = 0
                            End If
                            ' on ajout des lignes à partir du mois suivant , jusqu'au mois précédent le mvt de stock en cours
                            ajouteLignesRecapMensuelMoisVide(pDS, objFRN, getDateDebutdeMoisSuivant(dDatePrec), getDateFindeMoisPrecedent(oMvt.STK_DATE), nSF, pCout)
                        End If

                        'Traitement d'une ligne de mouvement de stock
                        '======================================================================

                        dDatePrec = oMvt.STK_DATE
                        bTrtCommence = True
                        'Si c'un Mouvement d'inventaire
                        If oMvt.STK_TYPE = vncTypeMvt.vncMvtInventaire Then
                            nSI = nSI + oMvt.STK_QTECOLIS
                            nSF = nSI
                        Else
                            If oMvt.STK_QTECOLIS > 0 Then
                                nEntree = oMvt.STK_QTECOLIS
                                nSortie = 0
                            End If
                            If oMvt.STK_QTECOLIS < 0 Then
                                nSortie = oMvt.STK_QTECOLIS
                                nEntree = 0
                            End If
                            If oMvt.STK_QTECOLIS = 0 Then
                                nSortie = 0
                                nEntree = 0
                            End If

                            nSF = nSI + oMvt.STK_QTECOLIS 'Le Stock Final est égal au Stock initial + Mvt Stock
                        End If

                        'Creation d'une Ligne de Recap
                        If (nSI <> 0 Or nEntree <> 0 Or nSortie <> 0 Or nSF <> 0) Then
                            oRow = pDS.RECAP_COLISAGE.NewRECAP_COLISAGERow()
                            oRow.RC_TYPE = oMvt.STK_TYPE
                            oRow.RC_SI = nSI
                            oRow.RC_ENTREE = nEntree
                            oRow.RC_SORTIE = nSortie
                            oRow.RC_SF = nSF
                            oRow.RC_PRD_CODE = oMvt.PRD_CODE
                            oRow.RC_PRD_LIBELLE = oMvt.PRD_LIBELLE
                            oRow.RC_FRN_CODE = oMvt.FRN_CODE
                            oRow.RC_FRN_NOM = oMvt.FRN_NOM
                            oRow.RC_FRN_RS = oMvt.FRN_RS
                            oRow.RC_DATE = oMvt.STK_DATE
                            oRow.RC_LIBELLE = oMvt.STK_LIB
                            oRow.RC_COUT = 0
                            pDS.RECAP_COLISAGE.AddRECAP_COLISAGERow(oRow)
                        End If
                        nEntreeMois = nEntreeMois + nEntree
                        nSortieMois = nSortieMois + nSortie



                        'Le Stock final devient le Stock initial pour le mouvement suivant
                        nSI = nSF

                    Next
                    '=== Traitement des lignes jusqu'à la date finale
                    If bTrtCommence Then
                        '==== Solde du mois en cours
                        AjoutLigneRecapMoisEnCours(pDS, objFRN, pCout, nEntreeMois, nSortieMois)
                        '==== Boucle sur les mois jusqu'à la date finale
                        ajouteLignesRecapMensuelMoisVide(pDS, objFRN, getDateDebutdeMoisSuivant(dDatePrec), getDateFindeMois(pdFin), nSF, pCout)

                    End If
                End If
            Next 'oFourn
        Catch ex As Exception
            Debug.Assert(False, ex.Message, ex.StackTrace)
        End Try


        Return pDS

    End Function

    Private Shared Sub ajouteLigneRecapMensuel(ByVal pDS As dsVinicom, ByVal pDate As Date, ByVal pobjFRN As Fournisseur, ByVal pCout As Decimal, ByVal pSI As Decimal, ByVal pSF As Decimal, ByVal pEntree As Decimal, ByVal pSorties As Decimal)
        Dim oRow As dsVinicom.RECAP_COLISAGERow
        Dim tabMois As String() = {"", "Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"}

        If (pSI <> 0 Or pEntree <> 0 Or pSorties <> 0 Or pSF <> 0) Then
            oRow = pDS.RECAP_COLISAGE.NewRECAP_COLISAGERow()
            oRow.RC_TYPE = "9" ' RECAP MENSUEL
            oRow.RC_SI = pSI
            oRow.RC_ENTREE = pEntree
            oRow.RC_SORTIE = pSorties
            oRow.RC_SF = pSF
            oRow.RC_PRD_CODE = ""
            oRow.RC_PRD_LIBELLE = ""
            oRow.RC_FRN_CODE = pobjFRN.code
            oRow.RC_FRN_NOM = pobjFRN.nom
            oRow.RC_FRN_RS = pobjFRN.rs
            oRow.RC_DATE = New Date(pDate.Year, pDate.Month, Date.DaysInMonth(pDate.Year, pDate.Month)) ''Dernier jour du mois
            oRow.RC_LIBELLE = tabMois(pDate.Month) + " " + pDate.Year.ToString()
            oRow.RC_COUT = pSF * pCout
            pDS.RECAP_COLISAGE.AddRECAP_COLISAGERow(oRow)
        End If

    End Sub

    Private Shared Sub AjoutLigneRecapMoisEnCours(ByVal pDs As dsVinicom, ByVal poFRN As Fournisseur, ByVal pCout As Decimal, ByVal pEntree As Decimal, ByVal pSortie As Decimal)
        Dim oRow As dsVinicom.RECAP_COLISAGERow
        Dim dDFinMois As Date
        Dim nSI As Decimal

        If pDs.RECAP_COLISAGE.Count > 0 Then
            oRow = pDs.RECAP_COLISAGE(pDs.RECAP_COLISAGE.Count - 1)
            If oRow.RC_FRN_CODE.Equals(poFRN.code) Then

                dDFinMois = New Date(oRow.RC_DATE.Year, oRow.RC_DATE.Month, Date.DaysInMonth(oRow.RC_DATE.Year, oRow.RC_DATE.Month))
                nSI = oRow.RC_SF - pEntree - pSortie
                Trace.WriteLine(poFRN.code + ":AjoutLigneRecapMoisEnCours->" + dDFinMois)
                ajouteLigneRecapMensuel(pDs, dDFinMois, poFRN, pCout, nSI, oRow.RC_SF, pEntree, pSortie)
            End If
        End If

    End Sub

    Private Shared Function IsRuptureMois(ByVal pDateNew As Date, ByVal pDateOld As Date) As Boolean
        Dim bReturn As Boolean
        Dim oDateMvt As Date = New Date(pDateNew.Year, pDateNew.Month, 1)
        'Si le 1 jour du mois du mvt est Sup à la date précédente => on a changé de mois
        bReturn = (pDateOld < oDateMvt)
        Return bReturn
    End Function
    Private Shared Function getDateFindeMois(ByVal pDate As Date) As Date
        Return New Date(pDate.Year, pDate.Month, Date.DaysInMonth(pDate.Year, pDate.Month))
    End Function
    Private Shared Function getDateDebutdeMois(ByVal pDate As Date) As Date
        Return New Date(pDate.Year, pDate.Month, 1)
    End Function
    Private Shared Function getDateDebutdeMoisSuivant(ByVal pDate As Date) As Date
        Return getDateDebutdeMois(pDate.AddMonths(1))
    End Function
    Private Shared Function getDateDebutdeMoisPrecedent(ByVal pDate As Date) As Date
        Return getDateDebutdeMois(pDate.AddMonths(-1))
    End Function
    Private Shared Function getDateFindeMoisPrecedent(ByVal pDate As Date) As Date
        Return getDateFindeMois(pDate.AddMonths(-1))
    End Function
    Private Shared Function getDateFindeMoisSuivant(ByVal pDate As Date) As Date
        Return getDateFindeMois(pDate.AddMonths(1))
    End Function
    Private Shared Function ajouteLignesRecapMensuelMoisVide(ByVal pDS As dsVinicom, ByVal poFRN As Fournisseur, ByVal pDateDeb As Date, ByVal pDateFin As Date, ByVal pSF As Integer, ByVal pCout As Decimal) As Boolean
        Dim dFin As Date
        Dim bReturn As Boolean
        Dim dDatePrec As Date = Date.MinValue
        'La Date limite est en fait le permier jour du mois suivant
        Dim dDateLimite As Date = getDateDebutdeMoisSuivant(pDateFin)
        Try
            Trace.WriteLine(poFRN.code + ":ajouteLignesRecapMensuelMoisVide->" + pDateDeb)
            While (getDateFindeMois(pDateDeb) < dDateLimite)
                '' Calcul du dernier jour du mois
                dFin = getDateFindeMois(pDateDeb)
                AjouteLigneRecapMensuel(pDS, dFin, poFRN, pCout, pSF, pSF, 0, 0)
                pDateDeb = pDateDeb.AddMonths(1) 'Mois Suivant
            End While
            bReturn = True
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn

    End Function

    Private Function AjouteLigneFactColisage(ByVal pDDeb As Date, ByVal pDFin As Date, ByVal pStockInitial As Decimal, ByVal pEntrees As Decimal, ByVal pSorties As Decimal, ByVal pStockFinal As Decimal) As Boolean
        Dim bReturn As Boolean
        Dim oLgFactcol As LgFactColisage

        Try
            oLgFactcol = New LgFactColisage()
            oLgFactcol.num = getNextNumLg()
            oLgFactcol.dDeb = pDDeb
            oLgFactcol.dFin = pDFin
            oLgFactcol.StockInitial = pStockInitial
            oLgFactcol.Entrees = pEntrees
            oLgFactcol.Sorties = pSorties
            oLgFactcol.StockFinal = pStockFinal
            oLgFactcol.qte = pStockFinal
            oLgFactcol.pu = CDec(Param.getConstante("CST_FACT_COL_PU_COLIS"))
            oLgFactcol.calculPrixTotal()

            AjouteLigneFactColisage(oLgFactcol)

            bReturn = True
        Catch ex As Exception
            setError(ex.StackTrace, ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function

    Public Overrides Sub Exporter(ByVal pstrFileName As String)
        MyBase.ExporterFacture(pstrFileName, "COL")
    End Sub

End Class
