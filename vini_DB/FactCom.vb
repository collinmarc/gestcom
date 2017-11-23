Imports System.IO
Imports System.Collections.Generic

Public Class FactCom
    Inherits Facture

    Private m_dateStat As Date
    Private m_periode As String
    Private m_montantReglement As Decimal
    Private m_refReglement As String
    Private m_dateReglement As Date


#Region "Accesseurs"
    Public Sub New(ByVal poTiers As Tiers)
        MyBase.New(poTiers, EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactComGeneree))
        Debug.Assert(poTiers.typeDonnee = vncEnums.vncTypeDonnee.FOURNISSEUR, "Tiers de type fournisseur requis")
        m_periode = ""
        m_montantReglement = 0
        m_refReglement = ""
        m_dateReglement = DATE_DEFAUT
        m_typedonnee = vncEnums.vncTypeDonnee.FACTCOMM
        m_dateStat = DATE_DEFAUT
        idModeReglement = poTiers.idModeReglement1 'utilisation du mode de reglement du Tiers
        'Pas de dupplication des commentaires car le commentaire facturation du client contient du commentaire 
        'pour le calcul du taux de commission
        'CommFacturation.comment = poTiers.CommFacturation.comment
    End Sub
    Public Shared Function createandload(ByVal pid As Long) As FactCom
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactCom
        obj = New FactCom
        Try
            If pid <> 0 Then
                obj.load(pid)
            End If
        Catch ex As Exception
            setError("CreateFactCom", ex.ToString)
        End Try
        Debug.Assert(obj.id = pid, "FactCom " & pid & " non chargée")
        Return obj
    End Function 'createandload
    Public Shared Function createandload(ByVal pCode As String) As FactCom
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactCom = Nothing
        Dim oCol As Collection
        Try
            obj = New FactCom()
            oCol = getListe(pCode)
            If (oCol.Count = 1) Then
                obj = oCol(1)
            End If

        Catch ex As Exception
            setError("CreateFactCom", ex.ToString)
        End Try
        Debug.Assert(obj.code = pCode, "FactCom " & pCode & " non chargée" + getErreur())
        Return obj
    End Function 'createandload
    Friend Sub New()
        MyClass.New(New Fournisseur)
    End Sub
    Public Property dateReglement() As Date
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
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
            Debug.Assert(Not m_bResume, "Objet de type resumé")
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
            Debug.Assert(Not m_bResume, "Objet de type resumé")
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
            Debug.Assert(Not m_bResume, "Objet de type resumé")
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
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_montantReglement
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_montantReglement Then
                RaiseUpdated()
                m_montantReglement = Value
            End If
        End Set

    End Property

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return code & " | " & TiersRS & " | " & dateCommande.ToShortDateString & " | " & Format(totalHT, "C") & " | " & Format(totalTTC, "C") & " | " & etat.libelle
        End Get
    End Property


#End Region

    Public Shared Function createFactComs(ByVal pcolSCMD As List(Of SousCommande), Optional ByVal pdateFact As Date = DATE_DEFAUT, Optional ByVal pdateStat As Date = DATE_DEFAUT, Optional ByVal pPeriode As String = "") As ColEvent

        '======================================================================================
        'Function : createFactComs
        'description : création des factures de commission à partir d'une collection de souscommande
        '               Pour chaque Fournisseur Différent une facture de commission est crée et les 
        '           sous-Commande de ce fournisseur lui sont attribué
        '======================================================================================
        Debug.Assert(Not pcolSCMD Is Nothing)
        Dim ocolReturn As ColEvent
        Dim objSCMD As SousCommande
        Dim objFactCom As FactCom

        Try
            'Traitement des paramètres
            If pdateFact = DATE_DEFAUT Then
                pdateFact = Now()
            End If
            If pdateStat = DATE_DEFAUT Then
                pdateFact = pdateFact
            End If
            If pPeriode = "" Then
                pdateFact = CStr(pdateStat)
            End If

            ocolReturn = New ColEvent
            'Parcours de toutes les sousCommandes de la collection
            For Each objSCMD In pcolSCMD
                If objSCMD.Selected Then

                    'A-t-on déjà crée une facture de commission pour ce fournissseur
                    If ocolReturn.keyExists(objSCMD.oFournisseur.code) Then
                        'Une Facture à déjà été créée pour ce fournisseur
                        objFactCom = ocolReturn(objSCMD.oFournisseur.code)
                    Else
                        'Il n'y a pas de facture pour ce fournisseur
                        'Création de la facture de commission
                        objSCMD.oFournisseur.load()
                        objFactCom = New FactCom(objSCMD.oFournisseur)
                        objFactCom.dateFacture = pdateFact
                        objFactCom.dateStatistique = pdateStat
                        objFactCom.periode = pPeriode
                        'Ajout de la facture à la collection 
                        ocolReturn.Add(objFactCom, objSCMD.oFournisseur.code)
                    End If
                    'Ajout de la SousCommande dans la facture
                    objFactCom.AjouteLigneFactCom(objSCMD)
                End If 'Selected

            Next
        Catch ex As Exception
            ocolReturn = Nothing
            setError("FactCom.createFactCom", ex.ToString())
        End Try

        Debug.Assert(Not ocolReturn Is Nothing, getErreur())
        Return ocolReturn

    End Function
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNomFournisseur As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = ListeFACTCOM(strCode, strNomFournisseur, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListe" & getErreur())
        shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListeNonReglee(Optional ByVal strCode As String = "", Optional ByVal strNomFournisseur As String = "") As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeFACTCOMNonReglee(strCode, strNomFournisseur)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListeNonReglee" & getErreur())
        shared_disconnect()
        Return colReturn
    End Function
    Public Shared Function getListe(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pCodeFourn As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeFACTCOMEtat(pddeb, pdfin, pCodeFourn, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function
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
        Dim objCommande As FactCom

        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj
            bReturn = bReturn And (m_periode.Equals(objCommande.periode))

            bReturn = bReturn And (m_montantReglement.Equals(objCommande.montantReglement))
            bReturn = bReturn And (m_dateReglement.ToShortDateString().Equals(objCommande.dateReglement.ToShortDateString()))
            bReturn = bReturn And (m_refReglement.Equals(objCommande.refReglement))
            bReturn = bReturn And (m_dateStat.ToShortDateString().Equals(objCommande.dateStatistique.ToShortDateString()))

            Return bReturn

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals
#End Region
#Region "Interface Persist"
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        Try
            If pid <> 0 Then
                m_id = pid
            End If

            Debug.Assert(id <> 0, "idFactCom <> 0")
            bReturn = loadFACTCOM()
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
                m_bcolLgLoaded = loadcolLignes()
            End If
            Return bReturn
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactCom.DBLoad: " & getErreur())
        Return bReturn
    End Function

    Friend Overrides Function delete() As Boolean

        Dim bReturn As Boolean
        Try
            bReturn = detachecolSCMD()
            If bReturn Then
                bReturn = deleteFACTCOM()
            End If
            deleteReglements()
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactCom.delete: " & getErreur())
        Return bReturn

    End Function


    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Fournisseur n'est pas Renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Fournisseur n'est pas sauvegardé")
        Debug.Assert(id = 0, "id= 0")

        Dim bReturn As Boolean

        Try
            bReturn = False
            If setNewcode() Then
                bReturn = insertFACTCOM()
                'Rattachement des sous-commande à la Facture de commission
                If bReturn Then
                    bReturn = attachecolSCMD()
                End If
            End If
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactCom.insert: " & getErreur())
        Return bReturn

    End Function 'insert


    Friend Overrides Function update() As Boolean
        '================================================================================
        'Fonction: update
        'Description : Mise à jour de la Facture de Comission en base
        '================================================================================
        Debug.Assert(Not oTiers Is Nothing, "Le Fournisseur n'est pas renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Fournisseur n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        Try
            'Mise à jour de la sous-commande
            bReturn = UpdateFACTCOM()
            'Rattachement des sous-commande à la Facture de commission
            If bReturn Then
                bReturn = attachecolSCMD()
            End If
            Return bReturn
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactCom.Update: " & getErreur())
        Return bReturn

    End Function

    Private Function attachecolSCMD() As Boolean
        '=============================================================================
        ' Function : attachecolSCMD
        ' Description : relie les lignes de commandes (= sous-Commande) à la facture de commission
        '=============================================================================
        Dim bReturn As Boolean = True
        Dim objLgCommande As LgCommande
        Dim idSCMD As Integer
        Dim objSCMD As SousCommande

        Try
            bReturn = True
            'Parcours de la collection des Lignes
            For Each objLgCommande In colLignes
                'Pour Chaque Ligne => Chargement de la Sous commande Correspondante
                idSCMD = objLgCommande.idSCmd
                objSCMD = New SousCommande
                objSCMD.load(idSCMD)
                If objSCMD.id <> 0 Then
                    'Changement d'état de la Sous Commande pour indiquer qu'elle est Facturée
                    objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)
                    objSCMD.idFactCom = Me.id
                    'Sauvegarde de la sous Commande si necessaire
                    bReturn = objSCMD.Save()
                    Debug.Assert(bReturn, SousCommande.getErreur())
                End If
            Next

        Catch ex As Exception
            bReturn = False
            setError("FactCom.attachecolSCMD", ex.ToString)

        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'attachecolSCMD
    Private Function detachecolSCMD() As Boolean
        '=============================================================================
        ' Function : detachecolSCMD
        ' Description : Supprime le lien entre les sous commande et la facture de Commission
        '=============================================================================
        Dim bReturn As Boolean = True
        Dim objLgCommande As LgCommande
        Dim idSCMD As Integer
        Dim objSCMD As SousCommande

        Try
            bReturn = True
            'Parcours de la collection des Lignes
            For Each objLgCommande In colLignes
                'Pour Chaque Ligne => Chargement de la Sous commande Correspondante
                idSCMD = objLgCommande.idSCmd
                objSCMD = New SousCommande
                objSCMD.load(idSCMD)
                If objSCMD.id <> 0 Then
                    'Changement d'état de la Sous Commande pour indiquer qu'elle est Facturée
                    objSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDAnnFacturer)
                    objSCMD.idFactCom = 0
                    'Sauvegarde de la sous Commande si necessaire
                    bReturn = objSCMD.Save()
                    Debug.Assert(bReturn, SousCommande.getErreur())
                End If
            Next

        Catch ex As Exception
            bReturn = False
            setError("FactCom.detachecolSCMD", ex.ToString)

        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'detachecolSCMD
    Public Overrides Function loadcolLignes() As Boolean
        Debug.Assert(m_id <> 0, "La commande doit être sauvegardé au Préalable")
        Dim bReturn As Boolean
        shared_connect()
        If m_bcolLgLoaded Then
            m_colLignes.clear()
        End If
        bReturn = LoadcolLgFactCom()

        Debug.Assert(bReturn, ColEvent.getErreur())
        m_bcolLgLoaded = bReturn ' Les lignes de commandes sont chargées
        m_bColLgInsertorDelete = False
        shared_disconnect()
        Return bReturn
    End Function 'loadColLignes
    Public Shadows Function savecolLignes(ByVal bInsertUpdate As Boolean) As Boolean
        Debug.Assert(True, "Méthode désactivée dans ce contexte")
        Return False
    End Function

    Public Function purge() As Boolean
        '================================================================================
        ' Fonction : purge 
        ' Description : Supression de la facture ET de tous ses élements dépendants
        '================================================================================
        Dim bReturn As Boolean
        Dim objLgCMD As LgCommande
        Dim objSCMD As SousCommande
        Try
            bReturn = True
            If Not etat.codeEtat = vncEnums.vncEtatCommande.vncFactComExportee Then
                bReturn = False
                Exit Function
            End If
            'Chargement des lignes de factures de comm (Sous-Commandes)
            If Not m_bcolLgLoaded Then
                bReturn = loadcolLignes()
            End If
            If bReturn Then
                For Each objLgCMD In m_colLignes
                    objSCMD = SousCommande.createandload(objLgCMD.idSCmd)
                    Debug.Assert(Not objSCMD Is Nothing)
                    'Suppression de la sous-commande
                    objSCMD.bDeleted = True
                    bReturn = objSCMD.Save()
                    'La Commande n'est pas supprimée par cette purge
                    'Voir purge des commandes
                Next
            End If
            If bReturn Then
                m_colLignes = New ColEvent   'Pour Empêcher le traitrement des Sous- Commandes
                Me.bDeleted = True
                bReturn = Save()
            End If

        Catch ex As Exception
            setError("FactCom.purge", ex.ToString())
            bReturn = False
        End Try
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
        ncode = GetNumeroFactureCommission()
        shared_disconnect()
        If ncode = -1 Then
            breturn = False
        Else
            code = str & CStr(ncode)
            breturn = True
        End If
        Return breturn
    End Function 'setnewCode
    Public Overloads Function AjouteLigne(ByVal pobjLgCMD As LgCommande, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "Méthode désactivée dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "Méthode désactivée dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overloads Function AjouteLigneFactCom(ByVal p_oSCMD As SousCommande, Optional ByVal pCalcul As Boolean = True) As LgCommande
        '=======================================================================
        'Fonction : AjouteLigneFactCom()
        'Description : Créé une ligne de commande et l'ajoute à la collection via AjouteLigne
        'Détails    :   Appelle la Fonction AjoutLigne a
        'Retour : une ligne de commande ou nothing si l'ajout échoue
        '=======================================================================
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not p_oSCMD Is Nothing)
        Debug.Assert(p_oSCMD.oFournisseur.id = Me.oTiers.id, "Les identifiants de fournisseurs doivent être égaux")

        Dim oLgCmd As LgCommande
        Dim strNum As String

        Try
            strNum = getNextNumLg()
            oLgCmd = New LgCommande(Me.id, p_oSCMD.id)
            oLgCmd.num = strNum
            oLgCmd.oProduit = Nothing
            oLgCmd.qteCommande = 1
            oLgCmd.prixU = p_oSCMD.MontantCommission
            oLgCmd.bGratuit = False
            oLgCmd.prixHT = p_oSCMD.MontantCommission
            oLgCmd.prixTTC = 0
            p_oSCMD.changeEtat(vncEnums.vncActionEtatCommande.vncActionSCMDFacturer)

            m_colLignes.Add(oLgCmd, CStr(p_oSCMD.code))
            If pCalcul Then
                calculPrixTotal()
            End If
            m_bcolLgLoaded = True
        Catch ex As Exception
            setError("AjouteLigneFactCom", ex.ToString)
            oLgCmd = Nothing
        End Try
        Debug.Assert(Not oLgCmd Is Nothing, getErreur())
        Return oLgCmd
    End Function 'AjouteLigneFactCom
    Public Overrides Function calculPrixTotal() As Boolean
        Dim oLg As LgCommande
        Dim nHT As Decimal
        Dim nTTC As Decimal
        Dim objParam As Param

        nHT = 0
        nTTC = 0
        For Each oLg In m_colLignes
            nHT = nHT + oLg.prixHT
        Next oLg

        objParam = Param.TVAdefaut
        nTTC = Math.Round(nHT * (1 + (objParam.valeur / 100)), 2)

        'On passe par les accesseurs pour lever l'event Updated s'il y a lieu
        totalHT = nHT
        totalTTC = nTTC
        Return True
    End Function 'calculPrixTotal
#End Region


    Public Overrides Sub Exporter(ByVal pstFileName As String)
        MyBase.ExporterFacture(pstFileName, "COM")
    End Sub
    ''' <summary>
    ''' Renvoie une chaine de caractère au format CSV pour export internet
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function toCSV() As String
        Debug.Assert(bcolLignesLoaded, "Les Lignes doivent être chargées au préalable")
        Dim strResult1Line As String
        Dim strResult As String

        strResult = ""
        strResult1Line = ""
        strResult1Line = strResult1Line & Trim(Me.id) & ";"
        strResult1Line = strResult1Line & Trim(Me.code) & ";"
        strResult1Line = strResult1Line & Trim(Format(Me.dateFacture, "ddMMyyyy")) & ";"
        strResult1Line = strResult1Line & Trim(Me.oTiers.code) & ";"
        strResult1Line = strResult1Line & Trim(Format(Me.totalHT, "n")) & ";"
        strResult1Line = strResult1Line & Trim(Format(Me.totalTTC, "n")) & ";"
        strResult1Line = strResult1Line & Trim(Me.CommFacturation.comment) & ";"

        strResult1Line = Replace(strResult1Line, vbCrLf, "--")
        strResult1Line = Replace(strResult1Line, vbCr, "-")
        strResult1Line = Replace(strResult1Line, vbLf, "-")
        strResult1Line = Replace(strResult1Line, vbNullChar, "-")
        strResult1Line = Replace(strResult1Line, vbTab, "-")
        strResult1Line = Replace(strResult1Line, vbBack, "-")
        strResult = strResult & strResult1Line & vbCrLf
        Return strResult

    End Function 'ToCSV

End Class
