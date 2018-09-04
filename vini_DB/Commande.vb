'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Commande    
' Description : Classe de Base des Commandes 
'           Classes filles : Commandes Clients / Bon Appro / FactCom / FactTRP
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
'           Code    : Code Commande
'           DateCommande
'           Etat    : Etat Commande
'           oTiers  : Tiers Concerné
'           m_ocolLignes : Collection des Lignes
'           m_oTotalHT : Total HT de la commande
'           m_oTotalTTC : Total TTC de la commande
'Private
'===================================================================================================================================
'Historique :
'============
'24/11/04 : CalculPrixTotal : Exclusion des lignes gratuites (par sécurité et pour pouvoir corriger un bug dans le calcul des lignes)
'11/11/05 : Ajout de la logistique pour quelle soit commune entre les commandes client et les bons d'appro
'===================================================================================================================================
Public MustInherit Class Commande
    Inherits Persist
    Implements IComparable(Of Commande)


    '=================== MEMBRES PRIVES ====================================
    Private m_code As String                  'code
    Private m_dateCommande As Date                      'Date de Commande
    Private m_etat As EtatCommande                      'Etat de la commande
    Private m_oTiers As Tiers                           'Tiers de la commande
    Protected WithEvents m_colLignes As ColEvent        'Collection des lignes de commandes
    Protected m_oLgCourante As LgCommande               ' Ligne de commande courante (utilisée dans le set property)
    Protected m_TotalHT As Decimal                     'Total HT de la commande
    Protected m_idParamTVA As Integer                     'Taux de TVA
    Protected m_TotalTTC As Decimal                    'Total TTC de la commande
    Protected WithEvents m_oTransporteur As Transporteur       'Adresse du transporteur
    Protected m_dateValidation As Date                            'Date de Validation
    Protected m_dateEnlevement As Date                            'Date d'enlèvement souhaitée
    Protected m_dateLivraison As Date                             'Date de Livraison 
    Protected m_RefLivraison As String 'Date de Livraison
    Private WithEvents m_oCaracteristiquesTiers As Tiers             'Collection d'adresses
    Protected m_bcolLgLoaded As Boolean
    Protected m_bColLgInsertorDelete As Boolean                       'La Collectiondes Lignes a subi un insertion ou une suppresssion
    Protected Shared m_bChargerColLignes As Boolean = True             'Chargement automatique des lignes de la commandes (ParDéfaut Oui)
    'Logistique
    Private m_qteColis As Decimal
    Private m_qtePalettesNonPreparees As Decimal
    Private m_qtePalettesPreparees As Decimal
    Private m_poids As Decimal
    Private m_puPalettesNonPreparees As Decimal
    Private m_puPalettesPreparees As Decimal
    Private m_MontantTransport As Decimal

    Protected m_LettreVoiture As String = String.Empty ' N° de lettre voirtuire
    Private m_CoutTransport As Decimal = 0 'Cout du transport
    Private m_RefFactTrp As String = String.Empty 'reference de la facture Transporteur

    Public MustOverride Sub Exporter(ByVal pstrFileName As String)

    Public MustOverride Function genereMvtStock() As Boolean
    Public MustOverride Function ControleMvtStock() As Collection
    Public MustOverride Function supprimeMvtStock() As Boolean
    Public Overridable Function regenereMvtStock() As Boolean

        '=========================================================================================
        'Function : regenereMvtStock
        'Description : Regeneration des mouvements de stocks
        '           Supprimme tous les mouvements de stocks de la commande
        '           Recréé tous les mouvements de stocks de la commande
        '=========================================================================================

        Debug.Assert(Not m_bResume, "Objet de type resumé")
        Dim bReturn As Boolean

        etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncSupprimer
        bReturn = supprimeMvtStock()
        Debug.Assert(bReturn, "CommandeClient.controledesMvtsStocks" & getErreur())
        If bReturn Then
            'Regénération des Mvts de stocks de la commande
            etat.actionMvtStock = vncEnums.vncGenererSupprimer.vncGenerer
            loadcolLignes()
            bReturn = genereMvtStock()
            Debug.Assert(bReturn, "frmControleStock.controledesMvtsStocks" & getErreur())
            releasecolLignes()
        End If

    End Function 'RegenereMvtStock

    Friend MustOverride Function setNewcode() As Boolean

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "(" & m_code & "," & m_dateCommande & "," & m_etat.toString() & "," & m_oTiers.code & "," & m_TotalHT & "," & m_TotalTTC & ")"
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As Commande
        Dim oLgCommande As LgCommande

        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj
            bReturn = bReturn And (m_code.Equals(objCommande.code))
            bReturn = bReturn And (m_dateCommande.ToShortDateString().Equals(objCommande.dateCommande.ToShortDateString()))

            If Not m_oTiers Is Nothing Then
                bReturn = bReturn And (m_oTiers.Equals(objCommande.oTiers))
            Else
                bReturn = objCommande.oTiers Is Nothing
            End If
            bReturn = bReturn And (m_oCaracteristiquesTiers.Equals(objCommande.caracteristiqueTiers))

            bReturn = bReturn And (m_etat.Equals(objCommande.etat))
            bReturn = bReturn And (m_idParamTVA.Equals(objCommande.idParamTVA))

            For Each oLgCommande In m_colLignes
                bReturn = bReturn And oLgCommande.Equals(objCommande.colLignes(Str(oLgCommande.num)))
            Next
            bReturn = bReturn And (m_TotalHT = objCommande.totalHT)
            bReturn = bReturn And (m_TotalTTC = objCommande.totalTTC)
            bReturn = bReturn And (m_oTransporteur.Equals(objCommande.oTransporteur))
            bReturn = bReturn And (m_dateValidation.ToShortDateString().Equals(objCommande.dateValidation.ToShortDateString()))
            bReturn = bReturn And (m_dateEnlevement.ToShortDateString().Equals(objCommande.dateEnlevement.ToShortDateString()))
            bReturn = bReturn And (m_dateLivraison.Equals(objCommande.dateLivraison))
            bReturn = bReturn And (m_RefLivraison.Equals(objCommande.refLivraison))
            bReturn = bReturn And (m_qteColis.Equals(objCommande.qteColis))
            bReturn = bReturn And (m_qtePalettesNonPreparees.Equals(objCommande.qtePalettesNonPreparees))
            bReturn = bReturn And (m_qtePalettesPreparees.Equals(objCommande.qtePalettesPreparees))
            bReturn = bReturn And (m_poids.Equals(objCommande.poids))
            bReturn = bReturn And (m_puPalettesNonPreparees.Equals(objCommande.puPalettesNonPreparees))
            bReturn = bReturn And (m_puPalettesPreparees.Equals(objCommande.puPalettesPreparees))
            bReturn = bReturn And (m_MontantTransport.Equals(objCommande.montantTransport))
            bReturn = bReturn And (m_LettreVoiture.Equals(objCommande.lettreVoiture))
            bReturn = bReturn And (m_CoutTransport.Equals(objCommande.coutTransport))
            bReturn = bReturn And (m_RefFactTrp.Equals(objCommande.refFactTRP))


        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals
    Public Sub New(ByVal poTiers As Tiers, ByVal poEtat As EtatCommande)
        Debug.Assert(Not poTiers Is Nothing, "Tiers Non renseigné")
        Debug.Assert(Not poEtat Is Nothing, "Etat Non renseigné")
        m_etat = poEtat            'Etat crée dans ls classes filles
        m_code = ""
        m_dateCommande = Now().ToShortDateString
        m_dateLivraison = DateAdd(DateInterval.Day, 1, m_dateCommande).ToShortDateString
        m_dateEnlevement = DateAdd(DateInterval.Day, 1, m_dateCommande).ToShortDateString
        m_dateValidation = DATE_DEFAUT
        m_oTiers = poTiers
        m_colLignes = New ColEvent
        m_TotalHT = 0
        m_bcolLgLoaded = True
        m_RefLivraison = ""
        m_idParamTVA = Param.TVAdefaut.id
        oTransporteur = New Transporteur
        m_oTransporteur.Dupplique(Transporteur.TransporteurDefault)
        'Initialisationdu transporteur / defaut
        m_qteColis = 0
        m_qtePalettesPreparees = 0
        m_qtePalettesNonPreparees = 0
        m_poids = 0
        m_puPalettesPreparees = Param.getConstante("CST_PU_PALL_PREP")
        m_puPalettesNonPreparees = Param.getConstante("CST_PU_PALL_NONPREP")
        m_MontantTransport = 0

        m_LettreVoiture = ""
        m_CoutTransport = 0
        m_RefFactTrp = ""

        m_RefLivraison = ""
        m_oCaracteristiquesTiers = New Tiers("", "")
        m_bColLgInsertorDelete = False
        m_oLgCourante = Nothing

        Me.Selected = True  'Selection pour la création de factures de commsions

    End Sub


#Region "Accesseurs"
    Public Property code() As String
        Get
            Return m_code
        End Get
        Set(ByVal Value As String)
            If (Not m_code.Equals(Value)) Then
                Debug.Assert(Len(m_code) = 0, "Impossible de modifier un code Commande existant")
                m_code = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Sub FTO_setCode(pCode As String)
        m_code = pCode
        RaiseUpdated()
    End Sub


    Public Property dateCommande() As Date
        Get
            Return m_dateCommande
        End Get
        Set(ByVal Value As Date)
            If Value <> m_dateCommande Then
                m_dateCommande = Value
                bUpdated = True
            End If
        End Set
    End Property
    Public Property etat() As EtatCommande
        Get
            Return m_etat
        End Get
        Set(ByVal Value As EtatCommande)
            If m_etat Is Nothing Then
                If Not Value Is Nothing Then
                    m_etat = Value
                    bUpdated = True
                End If
            Else
                If Not m_etat.Equals(Value) Then
                    m_etat = Value
                    bUpdated = True
                End If
            End If

        End Set
    End Property
    Public ReadOnly Property EtatLibelle() As String
        Get
            Return m_etat.libelle
        End Get
    End Property 'EtatLibelle
    Public ReadOnly Property EtatCode() As vncEnums.vncEtatCommande
        Get
            Return m_etat.codeEtat
        End Get
    End Property 'EtatCode
    Public Property oTiers() As Tiers
        Get
            Return m_oTiers
        End Get
        Private Set(ByVal Value As Tiers)
            setTiers(Value)
        End Set
    End Property ' oTiers
    Public Sub setTiers(Value As Tiers)
        If m_oTiers Is Nothing Then
            If Not Value Is Nothing Then
                m_oTiers = Value
                'on ne dupplique les informations de tiers que si l'original était vide
                DuppliqueCaracteristiqueTiers()
                RaiseUpdated()
            End If
        Else
            If (m_oTiers.id = 0) Then
                If Not Value Is Nothing Then
                    m_oTiers = Value
                    'on ne dupplique les informations de tiers que si l'original était vide
                    DuppliqueCaracteristiqueTiers()
                    RaiseUpdated()
                End If
            Else
                '                If Not oTiers.Equals(Value) Then
                m_oTiers = Value
                '                   RaiseUpdated()
                '          End If
            End If
        End If

    End Sub

    ''' <summary>
    ''' Rend la Raison sociale du Tiers 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property TiersRS() As String
        Get
            If oTiers IsNot Nothing Then
                Return m_oTiers.rs
            Else
                Return String.Empty
            End If
        End Get
    End Property 'TiersRS
    ''' <summary>
    ''' Rend le code du tiers
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property TiersCode() As String
        Get
            If oTiers IsNot Nothing Then
                Return m_oTiers.code
            Else
                Return String.Empty
            End If
        End Get
    End Property 'TiersRS

    Public Overridable ReadOnly Property FournisseurRS() As String
        Get
            Return ""
        End Get
    End Property
    Public Overridable ReadOnly Property FournisseurCode() As String
        Get
            Return ""
        End Get
    End Property
    Public ReadOnly Property colLignes() As ColEvent
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colLignes
        End Get
    End Property
    Public Property totalHT() As Decimal
        Get
            Return Math.Round(m_TotalHT, 2)
        End Get
        Set(ByVal Value As Decimal)
            If m_TotalHT <> Value Then
                m_TotalHT = Math.Round(Value, 2)
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property totalTTC() As Decimal
        Get
            Return Math.Round(m_TotalTTC, 2)
        End Get
        Set(ByVal Value As Decimal)
            If m_TotalTTC <> Value Then
                m_TotalTTC = Math.Round(Value, 2)
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Dim str As String
            str = m_code
            If Not oTiers Is Nothing Then
                str = str & " | " & oTiers.rs
            End If
            str = str & " | " & dateCommande & " | " & Format(totalHT, "C") & " | " & etat.libelle
            Return str
        End Get
    End Property
    Private Sub m_ocolLignes_Updated() Handles m_colLignes.Updated
        RaiseUpdated()
    End Sub

    Public Property oTransporteur() As Transporteur
        Get
            Return m_oTransporteur
        End Get
        Set(ByVal Value As Transporteur)
            'If m_oTransporteur Is Nothing Then
            '    m_oTransporteur.load(Value.id)
            '    RaiseUpdated()
            'End If
            'If Not m_oTransporteur.Equals(Value) Then
            '    If Not Value Is Nothing Then
            '        m_oTransporteur.load(Value.id)
            '    Else
            '        m_oTransporteur = Nothing
            '    End If
            '    RaiseUpdated()
            'End If
            If m_oTransporteur Is Nothing Then
                If Not Value Is Nothing Then
                    m_oTransporteur = Value
                    RaiseUpdated()
                End If
            Else
                If Not m_oTransporteur.Equals(Value) Then
                    m_oTransporteur = Value
                    RaiseUpdated()
                End If
            End If
        End Set
    End Property
    Public ReadOnly Property TransporteurNom() As String
        Get
            Return m_oTransporteur.AdresseLivraison.nom
        End Get
    End Property

    Public Property idParamTVA() As Integer
        Get
            Return m_idParamTVA
        End Get
        Set(ByVal Value As Integer)
            If m_idParamTVA <> Value Then
                m_idParamTVA = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property dateValidation() As Date
        Get
            Return m_dateValidation
        End Get
        Set(ByVal Value As Date)
            If m_dateValidation <> Value Then
                m_dateValidation = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property dateLivraison() As Date
        Get
            Return m_dateLivraison
        End Get
        Set(ByVal Value As Date)
            If m_dateLivraison <> Value Then
                m_dateLivraison = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property dateEnlevement() As Date
        Get
            Return m_dateEnlevement
        End Get
        Set(ByVal Value As Date)
            If m_dateEnlevement <> Value Then
                m_dateEnlevement = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property refLivraison() As String
        Get
            Return m_RefLivraison
        End Get
        Set(ByVal Value As String)
            If Not m_RefLivraison.Equals(Value) Then
                m_RefLivraison = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public ReadOnly Property caracteristiqueTiers() As Tiers
        Get
            Return m_oCaracteristiquesTiers
        End Get
    End Property
    'Commentaire de commande
    Public ReadOnly Property CommCommande() As Commentaire
        Get
            Return m_oCaracteristiquesTiers.CommCommande
        End Get
    End Property
    Public Property CommentaireCommandeText() As String
        Get
            Return m_oCaracteristiquesTiers.CommCommande.comment
        End Get
        Set(ByVal value As String)
            m_oCaracteristiquesTiers.CommCommande.comment = value
        End Set
    End Property
    'Commentaire de commande
    Public ReadOnly Property CommLivraison() As Commentaire
        Get
            Return m_oCaracteristiquesTiers.CommLivraison
        End Get
    End Property
    Public Property CommentaireLivraisonText() As String
        Get
            Return m_oCaracteristiquesTiers.CommLivraison.comment
        End Get
        Set(ByVal value As String)
            m_oCaracteristiquesTiers.CommLivraison.comment = value
        End Set
    End Property
    'Commentaire de commande
    Public ReadOnly Property CommFacturation() As Commentaire
        Get
            Return m_oCaracteristiquesTiers.CommFacturation
        End Get
    End Property
    Public Property CommentaireFacturationText() As String
        Get
            Return m_oCaracteristiquesTiers.CommFacturation.comment
        End Get
        Set(ByVal value As String)
            m_oCaracteristiquesTiers.CommFacturation.comment = value
        End Set
    End Property
    'Commentaire libre
    Public ReadOnly Property CommLibre() As Commentaire
        Get
            Return m_oCaracteristiquesTiers.CommLibre
        End Get
    End Property

    Public Property CommentaireLibreText() As String
        Get
            Return m_oCaracteristiquesTiers.CommLibre.comment
        End Get
        Set(ByVal value As String)
            m_oCaracteristiquesTiers.CommLibre.comment = value
        End Set
    End Property

    Public ReadOnly Property bcolLignesLoaded() As Boolean
        Get
            Return m_bcolLgLoaded
        End Get
    End Property
    Public ReadOnly Property bcolLignesUpated() As Boolean
        Get
            Return m_bColLgInsertorDelete
        End Get
    End Property
    Public Shared Property bChargerColLignes As Boolean
        Get
            Return m_bChargerColLignes
        End Get
        Set(value As Boolean)
            m_bChargerColLignes = value
        End Set
    End Property

    Public Property qteColis() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_qteColis
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_qteColis Then
                m_qteColis = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qtePalettesPreparees() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_qtePalettesPreparees
        End Get
        Set(ByVal Value As Decimal)
            Trace.WriteLine("SetqtePalettesPreparees")
            If Value <> m_qtePalettesPreparees Then
                m_qtePalettesPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qtePalettesNonPreparees() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_qtePalettesNonPreparees
        End Get
        Set(ByVal Value As Decimal)
            Trace.WriteLine("SetqtePalettesNonPreparees")
            If Value <> m_qtePalettesNonPreparees Then
                m_qtePalettesNonPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property poids() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_poids
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_poids Then
                m_poids = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property puPalettesPreparees() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_puPalettesPreparees
        End Get
        Set(ByVal Value As Decimal)
            Trace.WriteLine("SetpuPalettesPreparees")
            If Value <> m_puPalettesPreparees Then
                m_puPalettesPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property puPalettesNonPreparees() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_puPalettesNonPreparees
        End Get
        Set(ByVal Value As Decimal)
            Trace.WriteLine("SetpuPalettesNonPreparees")
            If Value <> m_puPalettesNonPreparees Then
                m_puPalettesNonPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property montantTransport() As Decimal
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_MontantTransport
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_MontantTransport Then
                m_MontantTransport = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property coutTransport() As Decimal
        Get
            Return m_CoutTransport
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_CoutTransport Then
                m_CoutTransport = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property lettreVoiture() As String
        Get
            Return m_LettreVoiture
        End Get
        Set(ByVal Value As String)
            If Not m_LettreVoiture.Equals(Value) Then
                m_LettreVoiture = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property refFactTRP() As String
        Get
            Return m_RefFactTrp
        End Get
        Set(ByVal Value As String)
            If Not m_RefFactTrp.Equals(Value) Then
                m_RefFactTrp = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Rend le montant total de la Commssion sur la commande 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>valeur calculée = Somme des Commssions dur les lignes</remarks>
    Public ReadOnly Property montantCommission() As Decimal
        Get
            Dim nReturn As Decimal
            nReturn = 0
            For Each objLg As LgCommande In Me.colLignes
                nReturn = nReturn + objLg.MtComm
            Next
            Return nReturn
        End Get
    End Property
#End Region
    Friend Sub resetCode()
        m_code = ""
    End Sub
    '=======================================================================
    'Fonction : getNextNumLg()
    'Description : Rend le prochain Numéro de Ligne
    'Détails    :  
    'Retour : une Numéro de ligne
    '=======================================================================
    Public Overridable Function getNextNumLg() As Integer
        Dim oLg As LgCommande
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
    '=======================================================================
    'Fonction : loadcolLignes()
    'Description : Charge la collection des lignes
    'Détails    :  
    'Retour : Boolean
    '=======================================================================
    Public Overridable Function loadcolLignes() As Boolean
        Debug.Assert(m_id <> 0, "La commande doit être sauvegardé au Préalable")
        Dim bReturn As Boolean
        shared_connect()
        m_colLignes = New ColEvent
        bReturn = LoadcolLGCMD()
        For Each olg As LgCommande In colLignes
            olg.idTiers = oTiers.id
        Next
        m_bColLgInsertorDelete = False 'L'indicateur d'jout/suppr repasse à faux après le chargement car il a utilisé la méthode ajouteligne
        Debug.Assert(bReturn, racine.getErreur())
        m_bcolLgLoaded = bReturn ' Les lignes de commandes sont chargées
        shared_disconnect()
        Return bReturn
    End Function 'loadColLignes
    '=======================================================================
    'Fonction : releasecolLignes()
    'Description : DeCharge la collection des lignes
    'Détails    :  
    'Retour : Boolean
    '=======================================================================
    Public Function releasecolLignes() As Boolean
        m_colLignes = New ColEvent
        m_bcolLgLoaded = False ' Les lignes de commandes ne sont pas chargées
        m_bColLgInsertorDelete = False
        Return True
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
    Public Overridable Function savecolLignes() As Boolean
        Dim bReturn As Boolean

        Debug.Assert(Not m_colLignes Is Nothing, "m_col <> nothing")
        Debug.Assert(m_bcolLgLoaded, "La collection  doit être chargée au préalable")
        Debug.Assert(m_id <> 0, "La commande  doit être sauvegardée au préalable")
        'En mode commandeClient ilfaut supprimer les lignes avant de les recréer pour gérer les suppressions et ajouts de lignes.
        If m_bColLgInsertorDelete Then
            'On Supprime la collection avant de la recréer (Commande)
            bReturn = deletecolLgCMD()
            If bReturn Then
                bReturn = INSERTcolLGCMD()
            End If
            m_bColLgInsertorDelete = False
        Else
            'On Met à jour la collection (SousCommande)
            bReturn = UPDATEcolLGCMD()
        End If
        Debug.Assert(bReturn, "Commande.savecolLignes:" & getErreur())
        Return bReturn
    End Function
    'Methode : EstEntierementLivree
    'Description : rend Vrai si tous les lignes on été Livrées
    Public Function estEntierementLivree() As Boolean
        Dim olgCom As LgCommande
        Dim bReturn As Boolean

        bReturn = True
        For Each olgCom In colLignes
            bReturn = bReturn And olgCom.estLivree()
        Next

        Return bReturn

    End Function 'estEntierementLivree

    'Methode : EstPartiellementLivree
    'Description : rend Vrai si au moins une lignes est Livrée
    Public Function estPartiellementLivree() As Boolean
        Dim olgCom As LgCommande
        Dim bReturn As Boolean

        bReturn = False
        For Each olgCom In colLignes
            If olgCom.estLivree() Then
                bReturn = True
            End If
        Next

        Return bReturn

    End Function 'estEntierementLivree
    '=======================================================================
    'Fonction : AjouteLigne()
    'Description : Ajoute une ligne sur une commandeClient
    'Détails    :  
    'Retour : une ligne de commande ou nothing si l'ajout échoue
    '=======================================================================
    '    Public Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrixCommande As Boolean = True) As LgCommande
    Public Overloads Function AjouteLigne(ByVal pobjLgCMD As LgCommande, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not pobjLgCMD Is Nothing)
        Dim oReturn As LgCommande

        Try

            'Ajout du tiers dans la Ligne de commmande (utile pour touver le Tx de commission)
            pobjLgCMD.idTiers = m_oTiers.id

            If p_bCalculPrix Then
                pobjLgCMD.calculPrixTotal()
            End If
            m_colLignes.Add(pobjLgCMD, CStr(pobjLgCMD.num))
            If p_bCalculPrix Then
                calculPrixTotal()
            End If
            oReturn = pobjLgCMD
            m_bcolLgLoaded = True
            m_bColLgInsertorDelete = True
        Catch ex As Exception
            setError("Commande.AjouteLigne", "Ajout de ligne impossible key = " & pobjLgCMD.num)
            oReturn = Nothing
        End Try
        '        Debug.Assert(m_ocolLignes.Count = n + 1, "Le nombre d'élement dans la collection est incrémenté de 1")
        Return oReturn
    End Function 'AjouteLigne

    '=======================================================================
    'Fonction : AjouteLigne()
    'Description : Créé une ligne de commande et l'ajoute à la collection via AjouteLigne
    'Détails    :   Appelle la Fonction AjoutLigne a
    'Retour : une ligne de commande ou nothing si l'ajout échoue
    '=======================================================================
    Public Overloads Function AjouteLigne(ByVal p_strNum As String, ByVal p_oProduit As Produit, ByVal p_qteCmd As Decimal, ByVal p_prixU As Decimal, Optional ByVal p_bGratuit As Boolean = False, Optional ByVal p_prixHT As Decimal = -1, Optional ByVal p_prixTTC As Decimal = -1, Optional ByVal p_bCalculPrix As Boolean = True) As LgCommande
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(Not p_oProduit Is Nothing)

        Dim oLgCmd As LgCommande

        If m_typedonnee = vncTypeDonnee.BA Then
            oLgCmd = New LgCommande(0, , m_id)
        Else
            oLgCmd = New LgCommande(m_id)
        End If
        oLgCmd.num = p_strNum
        oLgCmd.idCmd = id
        oLgCmd.idSCmd = 0 ' La sous commande n'est pas connue
        oLgCmd.oProduit = p_oProduit
        oLgCmd.qteCommande = p_qteCmd
        oLgCmd.prixU = p_prixU
        oLgCmd.bGratuit = p_bGratuit
        oLgCmd.prixHT = p_prixHT
        oLgCmd.prixTTC = p_prixTTC

        oLgCmd = AjouteLigne(oLgCmd, p_bCalculPrix)
        Return oLgCmd
    End Function 'AjouteLigne

    '=======================================================================
    'Fonction : supprimeLigne()
    'Description : Supprime une ligne sur une commande
    'Détails    :  Le Numéro passé est le numéro de ligne et non l'indice dans la collection
    '
    'Retour : une ligne de commande ou nothing si l'ajout échoue
    '=======================================================================
    Public Overridable Function supprimeLigne(ByVal strNumLigne As String, Optional ByVal p_bCalculPrix As Boolean = True) As Boolean
        Debug.Assert(Not m_colLignes Is Nothing)
        Debug.Assert(m_bcolLgLoaded, "La collection des lignes doit être chargée")
        Dim bReturn As Boolean
        Dim objLg As LgCommande

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
                setError("Commande.SupprimeLigne(" & strNumLigne & ")", "Ligne inconnue")
                bReturn = False
            End If
        Catch ex As Exception
            setError("Commande.SupprimeLigne(" & strNumLigne & ")", ex.ToString)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SupprimeLigne

    '=======================================================================
    'Fonction : setEtat()
    'Description : Met a jour l'état de la commande
    'Détails    :  
    'Retour : Boolean
    '=======================================================================
    Public Function setEtat(ByVal nEtat As vncEnums.vncEtatCommande) As Boolean
        m_etat = EtatCommande.createEtat(nEtat)
        Return True
    End Function 'setEtat
    ''' <summary>
    ''' Calcul du prix total de la Commande + Commssion 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>le calcul de la commssion n'est effectuté que pour les commande</remarks>
    Public Overridable Function calculPrixTotal() As Boolean
        Dim oLg As LgCommande
        Dim nHT As Decimal
        Dim nTTC As Decimal


        nHT = 0
        nTTC = 0
        For Each oLg In m_colLignes
            If Not oLg.bGratuit Then
                nHT = nHT + oLg.prixHT
                nTTC = nTTC + oLg.prixTTC
            End If

        Next oLg
        CalcPoidsColis()
        'On passe par les accesseurs pour lever l'event Updated s'il y a lieu
        totalHT = nHT
        totalTTC = nTTC
        Return True
    End Function 'calculPrixTotal

    Public Function CalcPoidsColis() As Boolean
        Debug.Assert(m_bcolLgLoaded, "Les lignes doivent être chargées")

        Dim objLg As LgCommande
        Dim npoids As Decimal
        Dim nqteColis As Decimal
        Dim bReturn As Boolean

        Try

            npoids = 0
            nqteColis = 0
            For Each objLg In m_colLignes
                objLg.poids = 0
                objLg.qteColis = 0
                'On ne calcule le poids et le nbre de clis que sur les lignes payantes
                If Not objLg.bGratuit Then
                    Dim nqteGratuit As Decimal = 0
                    'REcherche des produits gratuits du même produit
                    For Each objLgG As LgCommande In m_colLignes
                        If objLgG.bGratuit And objLgG.oProduit.id = objLg.oProduit.id Then
                            If objLgG.qteLiv <> 0 Then
                                nqteGratuit = nqteGratuit + objLgG.qteLiv
                            Else
                                nqteGratuit = nqteGratuit + objLgG.qteCommande
                            End If
                        End If
                    Next
                    Dim nQteProduit As Decimal = 0
                    If objLg.qteLiv <> 0 Then
                        nQteProduit = objLg.qteLiv + nqteGratuit
                    Else
                        nQteProduit = objLg.qteCommande + nqteGratuit
                    End If

                    objLg.calcPoidsColis(nQteProduit)
                    npoids = npoids + objLg.poids
                    nqteColis = nqteColis + objLg.qteColis
                End If
            Next
            poids = npoids
            qteColis = nqteColis
            bReturn = True
        Catch ex As Exception
            setError("Commande.calcPoidsColis", ex.ToString)
            bReturn = False
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function

    Public Function CalcMontantTransport() As Boolean
        Dim bReturn As Boolean
        Try
            Trace.WriteLine("CalcMontantTransport")
            montantTransport = (qtePalettesPreparees * puPalettesPreparees) + (qtePalettesNonPreparees * puPalettesNonPreparees)
            bReturn = True
        Catch ex As Exception
            setError("CommandeClient.CalcMontantTransport", ex.ToString)
        End Try
        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function

    Public Overridable Function changeEtat(ByVal p_action As vncActionEtatCommande) As Boolean
        m_etat = m_etat.changeEtat(p_action)
        RaiseUpdated()
    End Function
    'Methode : getActionMvtStock
    'Description : renvoie l'action à réaliser en ce qui concerne les Mouvements de Stocks
    Public Overridable Function getActionMvtStock() As vncGenererSupprimer
        Return m_etat.actionMvtStock
    End Function 'getActionMvtStock
    'Methode : getActionSousCommande
    'Description : renvoie l'action à réaliser en ce qui concerne les SousCommandes
    Public Overridable Function getActionSousCommande() As vncGenererSupprimer
        Return m_etat.actionSousCommande
    End Function 'getActionSousCommande


    Private Sub m_oTransporteur_Updated() Handles m_oTransporteur.Updated
        RaiseUpdated()
    End Sub
    Public Overridable Function DuppliqueCaracteristiqueTiers() As Boolean
        Debug.Assert(Not oTiers Is Nothing)
        caracteristiqueTiers.nom = oTiers.nom
        caracteristiqueTiers.rs = oTiers.rs
        caracteristiqueTiers.AdresseLivraison.nom = oTiers.AdresseLivraison.nom
        caracteristiqueTiers.AdresseLivraison.rue1 = oTiers.AdresseLivraison.rue1
        caracteristiqueTiers.AdresseLivraison.rue2 = oTiers.AdresseLivraison.rue2
        caracteristiqueTiers.AdresseLivraison.cp = oTiers.AdresseLivraison.cp
        caracteristiqueTiers.AdresseLivraison.ville = oTiers.AdresseLivraison.ville
        caracteristiqueTiers.AdresseLivraison.tel = oTiers.AdresseLivraison.tel
        caracteristiqueTiers.AdresseLivraison.fax = oTiers.AdresseLivraison.fax
        caracteristiqueTiers.AdresseLivraison.port = oTiers.AdresseLivraison.port
        caracteristiqueTiers.AdresseLivraison.Email = oTiers.AdresseLivraison.Email
        caracteristiqueTiers.AdresseFacturation.nom = oTiers.AdresseFacturation.nom
        caracteristiqueTiers.AdresseFacturation.rue1 = oTiers.AdresseFacturation.rue1
        caracteristiqueTiers.AdresseFacturation.rue2 = oTiers.AdresseFacturation.rue2
        caracteristiqueTiers.AdresseFacturation.cp = oTiers.AdresseFacturation.cp
        caracteristiqueTiers.AdresseFacturation.ville = oTiers.AdresseFacturation.ville
        caracteristiqueTiers.AdresseFacturation.tel = oTiers.AdresseFacturation.tel
        caracteristiqueTiers.AdresseFacturation.fax = oTiers.AdresseFacturation.fax
        caracteristiqueTiers.AdresseFacturation.port = oTiers.AdresseFacturation.port
        caracteristiqueTiers.AdresseFacturation.Email = oTiers.AdresseFacturation.Email
        caracteristiqueTiers.banque = oTiers.banque
        caracteristiqueTiers.rib1 = oTiers.rib1
        caracteristiqueTiers.rib2 = oTiers.rib2
        caracteristiqueTiers.rib3 = oTiers.rib3
        caracteristiqueTiers.rib4 = oTiers.rib4
        caracteristiqueTiers.idModeReglement = oTiers.idModeReglement
        caracteristiqueTiers.libModeReglement = oTiers.libModeReglement
        Return True
    End Function 'DuppliqueCaracteristiqueTiers

    Private Sub m_oCaracteristiquesTiers_Updated() Handles m_oCaracteristiquesTiers.Updated
        RaiseUpdated()
    End Sub
    Public Function setLgCourante(pLg As LgCommande) As Boolean
        Dim bReturn As Boolean
        Try
            m_oLgCourante = pLg
            bReturn = True
        Catch ex As Exception
            Trace.Write(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function



    ''' <summary>
    ''' Export d'une Sous commande au format CSV pour être importée Par le Logiciel QuadraFact
    ''' </summary>
    ''' <returns>Nom du fichier généré ou ""</returns>
    ''' <remarks></remarks>
    Public Function toCSVQuadraFact(pstrFile As String, pType As vncTypeExportQuadra) As String
        Debug.Assert(bcolLignesLoaded, "Les Lignes doivent être chargées au préalable")
        Dim objLgCommande As LgCommande
        Dim objProduit As Produit
        Dim bReturn As Boolean
        Dim oTA As vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter
        Dim oDT As vini_DB.dsVinicom.EXPORTPARAMDataTable
        Dim oRow As vini_DB.dsVinicom.EXPORTPARAMRow
        Dim nbChamps As Integer
        Dim nChamps As Integer
        Dim strLine As String
        Dim strValeur As String
        Dim strCodeExport As String = "QFACT"

        bReturn = False
        Try


            oTA = New vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter()
            oTA.Connection = Persist.oleDBConnection

            'Lecture des champs de la ligne 1
            oDT = oTA.GetDataBy_TYPE_NLIGNE(strCodeExport, 1)
            'Recupéation du nombre de champs
            nbChamps = oTA.getNbChamps(strCodeExport, 1)

            'Création de l'entête du fichier si nécessaire
            If Not System.IO.File.Exists(pstrFile) Then
                strLine = ""
                For nChamps = 1 To nbChamps
                    oRow = oDT.FindByEXP_TYPEEXP_NLIGNEEXP_NCHAMPS(strCodeExport, 1, nChamps)
                    If oRow IsNot Nothing Then
                        If oRow.EXP_LONGUEUR = 0 Then
                            strLine = strLine + Trim(oRow.EXP_VALEUR)
                        Else
                            strLine = strLine + Left(oRow.EXP_VALEUR + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                        End If
                    End If

                Next

                System.IO.File.AppendAllText(pstrFile, strLine & vbCrLf)
            End If

            'Ajout d'une ligne par Ligne de produit
            For Each objLgCommande In colLignes
                strLine = ""
                objProduit = Produit.createandload(objLgCommande.oProduit.id) 'Chargement du produit

                For nChamps = 1 To nbChamps
                    'Pour chaque Champs
                    oRow = oDT.FindByEXP_TYPEEXP_NLIGNEEXP_NCHAMPS(strCodeExport, 1, nChamps)
                    If oRow IsNot Nothing Then
                        strValeur = ""
                        If Trim(oRow.EXP_TYPECHAMPS).Equals("C") Then
                            'Exportation d'une Constante
                            strValeur = oRow.EXP_VALEUR
                        End If
                        If Trim(oRow.EXP_TYPECHAMPS).Equals("V") Then
                            'Exportation d'une Valeur
                            strValeur = getAttributeValue(oRow.EXP_VALEUR, objLgCommande, pType)
                        End If

                        'Si la longueur est égale à 0 => Trim
                        If oRow.EXP_LONGUEUR = 0 Then
                            strLine = strLine + Trim(strValeur)
                        Else
                            strLine = strLine + Left(strValeur + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                        End If
                    End If
                Next 'champs
                'Remplacement des caractères spéciaux
                strLine = Replace(strLine, vbCrLf, "--")
                strLine = Replace(strLine, vbCr, "-")
                strLine = Replace(strLine, vbLf, "-")
                strLine = Replace(strLine, vbNullChar, "-")
                strLine = Replace(strLine, vbTab, "-")
                strLine = Replace(strLine, vbBack, "-")


                'Ajout de la ligne dans le fichier
                System.IO.File.AppendAllText(pstrFile, strLine & vbCrLf)

            Next objLgCommande
            bReturn = True

        Catch ex As Exception
            Trace.WriteLine("SousCommande.toCSVQuadra ERR" & ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'ToCSVQuadra

    Public Function getAttributeValue(ByVal pstrAttributeName As String, pLgCommande As LgCommande, pType As vncTypeExportQuadra) As String
        Dim strReturn As String
        strReturn = String.Empty

        Try

            Select Case pstrAttributeName.ToUpper()
                Case "IDENTIFIANT"
                    If pType = vncTypeExportQuadra.vncExportBafClient Then
                        strReturn = Trim(Me.TiersCode)
                    Else
                        strReturn = Trim(Me.FournisseurCode)
                    End If

                Case "REFERENCE1"
                    If pType = vncTypeExportQuadra.vncExportBafClient Then
                        strReturn = Trim(Me.getCodeCommande())
                    Else
                        strReturn = Trim(Me.code) '' code de la sous Commande
                    End If
                Case "DATEPIECE"
                    strReturn = Trim(Format(Me.dateCommande, "yyMMdd"))
                Case "PIECEREGROUP"
                    If pType = vncTypeExportQuadra.vncExportBafClient Then
                        strReturn = Trim(Me.getCodeCommande())
                    Else
                        strReturn = Trim(Me.code) '' code de la sous Commande
                    End If

                    strReturn = strReturn.Replace("-", "")
                    If strReturn.Length > 12 Then
                        'Troncation à 12 car à gauche
                        strReturn = Right(strReturn, 12)
                    End If
                Case "DATEPEIECE"
                    strReturn = Trim(Format(Me.dateCommande, "yyMMdd"))
                Case "CODEARTICLE"
                    strReturn = Trim(pLgCommande.oProduit.code)
                Case "QUANTITE"
                    strReturn = Trim(pLgCommande.qteLiv)
                Case "PRIXHT"
                    If pLgCommande.qteLiv <> 0 Then
                        strReturn = Trim(pLgCommande.prixHT / pLgCommande.qteLiv)
                    Else
                        strReturn = "0"
                    End If
                Case "CODEDEPOT"
                    strReturn = My.MySettings.Default.CODEDEPOTQUADRA

            End Select
        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            strReturn = ""
        End Try

        Return strReturn
    End Function

    ''' <summary>
    ''' Valider l'export quadra : Changement d'état de la sousCommande
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function ValiderExportQuadra() As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Rend le code de la commande (Par defaut le code)
    ''' Surchargé dans Sous commande pour renvoyer codecommande
    ''' Surcharge utilisé dans l'exportQuadra
    ''' </summary>
    ''' <returns></returns>
    Public Overridable Function getCodeCommande() As String
        Return code
    End Function

    Private m_origine As String
    Public Property Origine() As String
        Get
            Return m_origine
        End Get
        Set(ByVal value As String)
            m_origine = value
        End Set
    End Property
    Private bSelected As Boolean
    Public Property Selected() As Boolean
        Get
            Return bSelected
        End Get
        Set(ByVal value As Boolean)
            bSelected = value
        End Set
    End Property

    Public Overrides Function checkForDelete() As Boolean

    End Function

    Protected Overrides Function DBLoad(Optional pid As Integer = 0) As Boolean

    End Function

    Friend Overrides Function delete() As Boolean

    End Function

    Friend Overrides Function insert() As Boolean

    End Function

    Friend Overrides Function update() As Boolean

    End Function

    Public Function CompareTo(other As Commande) As Integer Implements IComparable(Of Commande).CompareTo
        Return dateCommande.CompareTo(other.dateCommande)
    End Function
End Class ' Commande
