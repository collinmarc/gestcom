Imports System.configuration
Public Class FactHBV
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
        MyBase.New(poClient, EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactHBVGeneree))
        m_typedonnee = vncEnums.vncTypeDonnee.FACTHBV
        idModeReglement = poClient.idModeReglement1 'utilisation du mode de reglement du Tiers
    End Sub
    Public Sub New(ByVal poCmd As CommandeClient)
        MyBase.New(poCmd.oTiers, EtatCommande.createEtat(vncEnums.vncEtatCommande.vncFactHBVGeneree))
        m_typedonnee = vncEnums.vncTypeDonnee.FACTHBV
        idModeReglement = poCmd.oTiers.idModeReglement1 'utilisation du mode de reglement du Tiers
        dateCommande = poCmd.dateCommande
        CommFacturation.comment = poCmd.CommentaireFacturationText
        For Each oLg As LgCommande In poCmd.colLignes
            Dim olgFHBV As New LgFactHBV()
            olgFHBV.oProduit = oLg.oProduit
            olgFHBV.qteCommande = oLg.qteCommande
            olgFHBV.qteLiv = oLg.qteLiv
            olgFHBV.qteFact = oLg.qteFact
            olgFHBV.prixU = oLg.prixU
            olgFHBV.prixHT = oLg.prixHT
            olgFHBV.prixTTC = oLg.prixTTC
            olgFHBV.bGratuit = oLg.bGratuit

            AjouteLigneFactHBV(olgFHBV)

        Next
    End Sub

    Public Shared Function createandload(ByVal pid As Long) As FactHBV
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactHBV
        obj = New FactHBV
        Try
            If pid <> 0 Then
                obj.load(pid)
            End If
        Catch ex As Exception
            setError("FactHBV.CreateandLoad", ex.ToString)
        End Try
        Debug.Assert(obj.id = pid, "FactHBV " & pid & " non charg�e")
        Return obj
    End Function 'createandload
    Public Shared Function createandload(ByVal pCode As String) As FactHBV
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactHBV = Nothing
        Dim oCol As Collection
        Try
            obj = New FactHBV()
            oCol = getListe(pCode)
            If (oCol.Count = 1) Then
                obj = oCol(1)
            End If

        Catch ex As Exception
            setError("CreateFactHBV", ex.Message)
        End Try
        Debug.Assert(obj.code = pCode, "FactHBV " & pCode & " non charg�e")
        Return obj
    End Function 'createandload
    Friend Sub New()
        Me.New(New Client)
    End Sub
    Public Property dateReglement() As Date
        Get
            Debug.Assert(Not bResume, "Objet R�sum�")
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
            Debug.Assert(Not bResume, "Objet R�sum�")
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
            Debug.Assert(Not bResume, "Objet R�sum�")
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
            Debug.Assert(Not bResume, "Objet R�sum�")
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
            Debug.Assert(Not bResume, "Objet R�sum�")
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
            Debug.Assert(Not bResume, "Objet R�sum�")
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
#Region "M�thode de racine"

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'D�tails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "(" & code & "," & dateCommande & "," & periode & "," & oTiers.code & "," & m_TotalHT & "," & m_TotalTTC & ")"
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'�quivalence entre 2 objets
    'D�tails    :  
    'Retour : Vari si l'objet pass� en param�tre est equivalent � 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As FactHBV

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

            bReturn = loadFACTHBV()
            If id <> 0 Then
                'Chargement des caract�ristiques du Client
                Try
                    oTiers.load()
                Catch ex As Exception
                    bReturn = False
                    setError("oTiers.load", Tiers.getErreur())
                End Try
                'Chargement des lignes de Facture
                m_bcolLgLoaded = loadcolLignes()
            End If


        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Return bReturn
    End Function 'DBLoad

    Friend Overrides Function delete() As Boolean

        Dim bReturn As Boolean = False
        Try
            bReturn = True
            bReturn = deletecolLgHBV()
            If bReturn Then
                bReturn = deleteFACTHBV()
            End If
        Catch ex As Exception
            bReturn = False
            setError("", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactHBV.delete: " & getErreur())
        Return bReturn

    End Function 'delete

    Public Overrides Function genereMvtStock() As Boolean
        Return False

    End Function
    ''' <summary>
    ''' insertion d'une facture dans la base de donn�es
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseign�")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegard�")
        Debug.Assert(id = 0, "id= 0")

        Dim bReturn As Boolean = False

        Try
            bReturn = False
            If setNewcode() Then
                bReturn = insertFACTHBV()
            End If
            If bcolLignesLoaded Then
                bReturn = savecolLignes()
            End If
        Catch ex As Exception
            bReturn = False
            setError("FactHBV.insert", ex.ToString())
        End Try

        Return bReturn

    End Function 'insert

    Public Overrides Function supprimeMvtStock() As Boolean
        Return False
    End Function

    Friend Overrides Function update() As Boolean
        '================================================================================
        'Fonction: update
        'Description : Mise � jour de la Facture de Transport
        '================================================================================
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas renseign�")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegard�")
        Debug.Assert(id <> 0, "idFact <> 0")
        Dim bReturn As Boolean = False
        Try

            'Mise � jour de la sous-commande
            bReturn = UpdateFACTHBV()
            If bcolLignesLoaded Then
                bReturn = savecolLignes()
            End If

        Catch ex As Exception
            bReturn = False
            setError("FactHBV.Update", ex.ToString())
        End Try

        Debug.Assert(bReturn, "FactHBV.Update: " & getErreur())
        Return bReturn

    End Function

    Public Overrides Function loadcolLignes() As Boolean
        Debug.Assert(m_id <> 0, "La Facture doit �tre sauvegard� au Pr�alable")
        Dim bReturn As Boolean = False
        shared_connect()
        If m_bcolLgLoaded Then
            m_colLignes.clear()
        End If
        bReturn = LoadcolLgFactHBV()

        m_bcolLgLoaded = bReturn ' Les lignes sont charg�es
        m_bColLgInsertorDelete = False ' Mais ne sont pas Mise � jour
        shared_disconnect()
        Return bReturn

    End Function 'loadColLignes
    '=======================================================================
    'Fonction : savecolLignes()
    'Description : Sauvegarde la collection des lignes
    '               En fonction du param�tre bDeleteInsert
    '                   Suppression des lignes pour reinsertion (CommandeClient)
    '                   Update des lignes
    'D�tails    :  
    'Retour : Boolean
    '=======================================================================
    Public Overrides Function savecolLignes() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(Not m_colLignes Is Nothing, "m_col <> nothing")
        Debug.Assert(m_bcolLgLoaded, "La collection  doit �tre charg�e au pr�alable")
        Debug.Assert(m_id <> 0, "La commande  doit �tre sauvegard�e au pr�alable")

        bReturn = saveColLGFACTHBV()
        Return bReturn
    End Function

    Public Function purge() As Boolean
        '================================================================================
        ' Fonction : purge 
        ' Description : Supression de la facture ET de tous ses �lements d�pendants
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
        '            'La Commande n'est pas supprim�e par cette purge
        '            'Voir purge des commandes
        '        Next
        '    End If
        '    If bReturn Then
        '        m_colLignes = New ColEvent   'Pour Emp�cher le traitrement des Sous- Commandes
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
#Region "M�thodes de commande"
    Friend Overrides Function setNewcode() As Boolean
        Dim str As String
        Dim ncode As Integer
        Dim breturn As Boolean

        str = ""
        shared_connect()
        ncode = GetNumeroFactureHBV()
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
    'Description : Rend le prochain Num�ro de Ligne
    'D�tails    :  
    'Retour : une Num�ro de ligne
    '=======================================================================
    Public Overrides Function getNextNumLg() As Integer
        Dim oLg As LgFactHBV
        Dim num As Integer = 0
        Dim bOk As Boolean
        'Nombre d'�lement dans la collection + 10
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
        Debug.Assert(False, "M�thode d�sactiv�e dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(False, "M�thode d�sactiv�e dans ce contexte")
        Return Nothing
    End Function 'AjouteLigne
    Public Overrides Function calculPrixTotal() As Boolean
        Return True
    End Function 'calculPrixTotal

    Public Overrides Function changeEtat(ByVal p_action As vncActionEtatCommande) As Boolean
        Debug.Assert(p_action >= vncActionEtatCommande.vncActionMinFactHBV And p_action <= vncActionEtatCommande.vncActionMaxFactHBV)
        Dim bReturn As Boolean = False
        Try
            If p_action >= vncActionEtatCommande.vncActionMinFactHBV And p_action <= vncActionEtatCommande.vncActionMaxFactHBV Then
                MyBase.changeEtat(p_action)
                RaiseUpdated()
                bReturn = True
            Else
                setError("FactHBV.changeEtat", "Code Action incorrect :" + p_action)
                bReturn = False
            End If

        Catch ex As Exception
            setError("FactHBV.changeEtat", ex.ToString)
            bReturn = False
        End Try

        Return bReturn
    End Function 'ChangeEtat

    Public Overrides Function ControleMvtStock() As Microsoft.VisualBasic.Collection
        Return Nothing
    End Function


#End Region
    Public Shared Function createFactHBVs(ByVal pcolCMD As Collection, Optional ByVal pdateFact As Date = DATE_DEFAUT, Optional ByVal pdateStat As Date = DATE_DEFAUT, Optional ByVal pPeriode As String = "") As ColEvent

        '======================================================================================
        'Function : createFactHBVs
        'description : cr�ation des factures de Transport � partir d'une collection de commande
        '               Pour chaque Client Diff�rent une facture de transport est cr�e et les 
        '           Commande de ce client lui sont attribu�
        '======================================================================================
        Debug.Assert(Not pcolCMD Is Nothing)
        Dim ocolReturn As ColEvent
        Dim objCMD As CommandeClient
        Dim objFactHBV As FactHBV

        Try
            'Traitement des param�tres
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
                'A-t-on d�j� cr�e une facture de Transport pour ce Client
                If ocolReturn.keyExists(objCMD.oTiers.code) Then
                    'Une Facture � d�j� �t� cr��e pour ce Client
                    objFactHBV = ocolReturn(objCMD.oTiers.code)
                Else
                    'Il n'y a pas de facture pour ce client
                    'Cr�ation de la facture de transport
                    objFactHBV = New FactHBV(objCMD.oTiers)
                    objFactHBV.dateFacture = pdateFact
                    objFactHBV.dateStatistique = pdateStat
                    objFactHBV.periode = pPeriode
                    objFactHBV.montantTaxes = 0
                    'Ajout de la facture � la collection 
                    ocolReturn.Add(objFactHBV, objFactHBV.oTiers.code)
                End If
                'Ajout de la taxe d'enregistrement autant de fois qu'il y a de lignes
                objFactHBV.montantTaxes = objFactHBV.montantTaxes + Param.getConstante("CST_TAXES_HBV")
                'Ajout de la Commande dans la facture
                '                objFactHBV.AjouteLigneFactHBV(objCMD, True)
            Next
        Catch ex As Exception
            ocolReturn = Nothing
            setError("FactHBV.createFactHBV", ex.ToString())
        End Try

        Debug.Assert(Not ocolReturn Is Nothing, getErreur())
        Return ocolReturn

    End Function 'createFactsHBV
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNomClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = ListeFACTHBV(strCode, strNomClient, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactHBV.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

        Return colReturn
    End Function
    Public Shared Function getListe(ByVal pddeb As Date, ByVal pdfin As Date, Optional ByVal pCodeClient As String = "", Optional ByVal pEtat As vncEtatCommande = vncEnums.vncEtatCommande.vncRien) As Collection
        Dim colReturn As Collection

        shared_connect()
        colReturn = Persist.ListeFACTHBVDate(pddeb, pdfin, pCodeClient, pEtat)
        Debug.Assert(Not colReturn Is Nothing, "FactCom.getListe" & getErreur())
        shared_disconnect()
        Return colReturn

    End Function

    Public Function AjouteLigneFactHBV(ByVal p_oLgFactHBV As LgFactHBV, Optional ByVal pCalcul As Boolean = True) As LgFactHBV
        m_bColLgInsertorDelete = True
        colLignes.Add(p_oLgFactHBV)
        Return p_oLgFactHBV
    End Function 'AjouteLigneFactHBV
    '=======================================================================
    'Fonction : supprimeLigne()
    'Description : Supprime une ligne sur une commande
    'D�tails    :  Le Num�ro pass� est le num�ro de ligne et non l'indice dans la collection
    '
    'Retour : une ligne de commande ou nothing si l'ajout �choue
    '=======================================================================
    Public Overrides Function supprimeLigne(ByVal strNumLigne As String, Optional ByVal p_bCalculPrix As Boolean = True) As Boolean
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(m_bcolLgLoaded, "La collection des lignes doit �tre charg�e")
        Dim bReturn As Boolean
        Dim objLg As LgFactHBV

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
                setError("FactHBV.SupprimeLigne(" & strNumLigne & ")", "Ligne inconnue")
                bReturn = False
            End If
        Catch ex As Exception
            setError("FactHBV.SupprimeLigne(" & strNumLigne & ")", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SupprimeLigneHBV

    ''========================================================================
    ''fonction : exporter
    ''DEscription : Exporte la facture de transport dans un format MAXIMA
    ''DEtails    : 1facture =1 Ligne avec le prix Total
    ''
    ''Retour : Boolean
    ''=========================================================================
    'Public Function exporter(ByVal strFile As String) As Boolean

    '    Debug.Assert(strFile <> "", "Fichier inconnu")
    '    Debug.Assert(bcolLignesLoaded, "Les lignes doivent �tre charg�es")

    '    Dim nvcEntete As System.Collections.Specialized.NameValueCollection
    '    Dim bReturn As Boolean
    '    Dim nFile As Integer
    '    Dim objLg As LgFactHBV
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

    '        'Premi�re ligne d'adresse
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

    '        'Deuxi�me ligne d'adresse
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
    '        setError("FactHBV.exporter", ex.ToString())
    '    End Try
    '    Return bReturn
    'End Function
    'Private Function exporterAttribut(ByVal nFile As Integer, ByVal strAttribut As String, ByVal obj As Object) As Boolean
    '    Dim objFact As FactHBV = Nothing
    '    Dim objLgFact As LgFactHBV = nothing

    '    Try
    '        objFact = CType(obj, FactHBV)
    '    Catch ex As Exception
    '        Try
    '            objLgFact = CType(obj, LgFactHBV)
    '        Catch ex1 As Exception
    '            Debug.Assert(False, "Objet de type FactHBV ou LGFactHBV requis")
    '            setError("FactHBV.exporterAttribut", "Objet de type FactHBV ou LGFactHBV requis")
    '            Return False
    '        End Try


    '    End Try
    '    Dim bReturn As Boolean
    '    Try
    '        bReturn = True
    '        Select Case strAttribut
    '            Case "FHBV_ID"
    '                Print(nFile, objFact.id)
    '            Case "FHBV_CODE"
    '                Print(nFile, objFact.code)
    '            Case "FHBV_DATE"
    '                Print(nFile, objFact.dateCommande.ToShortDateString)
    '            Case "FHBV_CLT_ID"
    '                Print(nFile, objFact.oTiers.id)
    '            Case "FHBV_TOTAL_TAXES"
    '                Print(nFile, objFact.montantTaxes)
    '            Case "FHBV_TOTAL_HT"
    '                Print(nFile, objFact.totalHT)
    '            Case "FHBV_TOTAL_TTC"
    '                Print(nFile, objFact.totalTTC)
    '            Case "FHBV_IDMODEREGLEMENT"
    '                Print(nFile, objFact.idModeReglement)
    '            Case "FHBV_PERIODE"
    '                Print(nFile, objFact.periode)
    '            Case "FHBV_ETAT"
    '                Print(nFile, objFact.etat.codeEtat)
    '            Case "FHBV_MONTANT_REGLEMENT"
    '                Print(nFile, objFact.montantReglement)
    '            Case "FHBV_DATE_REGLEMENT"
    '                Print(nFile, objFact.dateReglement)
    '            Case "FHBV_REF_REGLEMENT"
    '                Print(nFile, objFact.refReglement)
    '            Case "FHBV_COM_FACT"
    '                Print(nFile, objFact.CommFacturation)
    '            Case "FHBV_DATE_REGLEMENT"
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

    '            Case "LGHBV_ID"
    '                Print(nFile, objLgFact.id)
    '            Case "LGHBV_NUM"
    '                Print(nFile, objLgFact.num)
    '            Case "LGHBV_CODEMAXIMA"
    '                'Attention le code MAXIMA est calcul�
    '                If (objLgFact.num = CST_LGFACTHBV_NUM_GO) Then
    '                    Print(nFile, ConfigurationManager.AppSettings.GetValues("EXPORT_CODE_GAZOLE")(0))
    '                Else
    '                    Print(nFile, ConfigurationManager.AppSettings.GetValues("EXPORT_CODE_TRANSPORT")(0))
    '                End If
    '            Case "LGHBV_FACTHBV_ID"
    '                Print(nFile, objLgFact.idFactHBV)
    '            Case "LGHBV_CMDCLT_ID"
    '                Print(nFile, objLgFact.idCmdCLT)
    '            Case "LGHBV_HBVEUR"
    '                Print(nFile, objLgFact.nomTransporteur)
    '            Case "LGHBV_DATE_LIV"
    '                Print(nFile, objLgFact.dateLivraison)
    '            Case "LGHBV_REF_LIV"
    '                Print(nFile, objLgFact.referenceLivraison)
    '            Case "LGHBV_DATE_CMD"
    '                Print(nFile, objLgFact.dateCommande)
    '            Case "LGHBV_REF_CMD"
    '                Print(nFile, objLgFact.refCommande)
    '            Case "LGHBV_QTE_COLIS"
    '                Print(nFile, objLgFact.qteColis)
    '            Case "LGHBV_QTE_POIDS"
    '                Print(nFile, objLgFact.poids)
    '            Case "LGHBV_QTE_PAL_NONPREP"
    '                Print(nFile, objLgFact.qtePalettesNonPreparees)
    '            Case "LGHBV_QTE_PAL_PREP"
    '                Print(nFile, objLgFact.qtePalettesPreparees)
    '            Case "LGHBV_PU_PAL_NONPREP"
    '                Print(nFile, objLgFact.puPalettesNonPreparees)
    '            Case "LGHBV_PU_PAL_PREP"
    '                Print(nFile, objLgFact.puPalettesPreparees)
    '            Case "LGHBV_MT_HT"
    '                Print(nFile, objLgFact.prixHT)
    '            Case "LGHBV_RESUME"
    '                Print(nFile, objLgFact.dateLivraison.ToShortDateString & " " & objLgFact.referenceLivraison & " " & objLgFact.nomTransporteur & " " & objLgFact.dateCommande.ToShortDateString & " " & objLgFact.refCommande)
    '            Case Else
    '                Print(nFile, strAttribut)
    '        End Select
    '        bReturn = True
    '    Catch ex As Exception
    '        bReturn = False
    '        setError("FactHBV.exporterAttribut", ex.ToString)
    '    End Try
    '    Debug.Assert(bReturn, getErreur())

    '    Return bReturn
    'End Function

    Public Overrides Sub Exporter(ByVal pstrFileName As String)

        MyBase.ExporterFacture(pstrFileName, "HBV")
    End Sub

End Class
