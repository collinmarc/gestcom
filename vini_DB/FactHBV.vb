Imports System.Configuration
Imports System.Collections.Generic

Public Class FactHBV
    Inherits Facture

    Private m_idCommande As Long

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
        m_idCommande = poCmd.id
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
    Public Property idCommande() As Long
        Get
            Return m_idCommande
        End Get
        Set(ByVal Value As Long)
            If Not m_idCommande.Equals(Value) Then
                m_idCommande = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public ReadOnly Property colLignesFactHBV As Collection
        Get
            Return colLignes.col
        End Get
    End Property


#End Region
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
        Debug.Assert(obj.id = pid, "FactHBV " & pid & " non chargée")
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
        Debug.Assert(obj.code = pCode, "FactHBV " & pCode & " non chargée")
        Return obj
    End Function 'createandload
    ''' <summary>
    ''' Creation et chargement d'une factute HBV à partir d'un id de commande
    ''' </summary>
    ''' <param name="pIdCmd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function createandloadFromCmd(ByVal pIdCmd As Long) As FactHBV
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim obj As FactHBV = Nothing
        Try
            obj = New FactHBV()
            obj.loadFACTHBVFromCmd(pIdCmd)
            obj.resetBooleans()

        Catch ex As Exception
            setError("createandloadFromCmd", ex.Message)
        End Try
        Return obj
    End Function 'createandload
    Private Sub New()
        Me.New(New Client)
    End Sub


    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return code & " | " & oTiers.rs & " | " & dateCommande.ToShortDateString & " | " & Format(totalTTC, "C") & " | " & etat.libelle
        End Get
    End Property

#Region "Méthode de racine"

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "(" & code & "," & dateCommande & "," & oTiers.code & "," & m_TotalHT & "," & m_TotalTTC & ")"
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As FactHBV

        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj

            bReturn = bReturn And (m_IDModeReglement.Equals(objCommande.idModeReglement))
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
                'Chargement des caractéristiques du Client
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
    Protected Function DBLoadFromCmd(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean = False
        Try
            If pid <> 0 Then
                m_id = pid
            End If

            bReturn = loadFACTHBV()
            If id <> 0 Then
                'Chargement des caractéristiques du Client
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
    ''' insertion d'une facture dans la base de données
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Overrides Function insert() As Boolean
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas Renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
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
        'Description : Mise à jour de la Facture de Transport
        '================================================================================
        Debug.Assert(Not oTiers Is Nothing, "Le Client n'est pas renseigné")
        Debug.Assert(oTiers.id <> 0, "Le Client n'est pas sauvegardé")
        Debug.Assert(id <> 0, "idFact <> 0")
        Dim bReturn As Boolean = False
        Try

            'Mise à jour de la sous-commande
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
        Debug.Assert(m_id <> 0, "La Facture doit être sauvegardé au Préalable")
        Dim bReturn As Boolean = False
        shared_connect()
        If m_bcolLgLoaded Then
            m_colLignes.clear()
        End If
        bReturn = LoadcolLgFactHBV()

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

        bReturn = saveColLGFACTHBV()
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
    'Description : Rend le prochain Numéro de Ligne
    'Détails    :  
    'Retour : une Numéro de ligne
    '=======================================================================
    Public Overrides Function getNextNumLg() As Integer
        Dim oLg As LgFactHBV
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
        totalHT = 0
        totalTTC = 0
        For Each oLg As LgFactHBV In colLignes
            '            oLg.calculPrixTotal()
            totalHT = totalHT + oLg.prixHT
            totalTTC = totalTTC + oLg.prixTTC
        Next
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
        If pCalcul Then
            calculPrixTotal()
        End If
        Return p_oLgFactHBV
    End Function 'AjouteLigneFactHBV
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
    '    Debug.Assert(bcolLignesLoaded, "Les lignes doivent être chargées")

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
    '                'Attention le code MAXIMA est calculé
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
