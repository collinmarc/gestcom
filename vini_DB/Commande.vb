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
        Set(ByVal Value As Tiers)
            If oTiers Is Nothing Or (oTiers IsNot Nothing And oTiers.id = 0) Then
                If Not Value Is Nothing Then
                    m_oTiers = Value
                    'on ne dupplique les informations de tiers que si l'original était vide
                    DuppliqueCaracteristiqueTiers()
                    RaiseUpdated()
                End If
            Else
                If Not oTiers.Equals(Value) Then
                    m_oTiers = Value
                    RaiseUpdated()
                End If
            End If
        End Set
    End Property ' oTiers
    ''' <summary>
    ''' Rend la Raison sociale du Tiers 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TiersRS() As String
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
    Public ReadOnly Property TiersCode() As String
        Get
            If oTiers IsNot Nothing Then
                Return m_oTiers.code
            Else
                Return String.Empty
            End If
        End Get
    End Property 'TiersRS
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
            bReturn = UpdatecolLgCMD()
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
    Protected Overloads Function AjouteLigne() As LgCommande
        Dim oLgCmd As LgCommande

        If m_typedonnee = vncTypeDonnee.BA Then
            oLgCmd = New LgCommande(0, , m_id)
        Else
            oLgCmd = New LgCommande(m_id)
        End If
        oLgCmd.num = getNextNumLg()
        oLgCmd = AjouteLigne(oLgCmd, False)
        Return oLgCmd

    End Function

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

            'Calcul de la commssion s'il y a lieu
            If m_typedonnee = vncTypeDonnee.COMMANDECLIENT Then
                If (Me.etat.codeEtat = vncEtatCommande.vncEnCoursSaisie Or Me.etat.codeEtat = vncEtatCommande.vncValidee) Then
                    oLg.CalculCommission(CalculCommQte.CALCUL_COMMISSION_QTE_CMDE)
                Else
                    oLg.CalculCommission(CalculCommQte.CALCUL_COMMISSION_QTE_LIVREE)
                End If
            End If
        Next oLg

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
                objLg.calcPoidsColis()
                npoids = npoids + objLg.poids
                nqteColis = nqteColis + objLg.qteColis
            Next
            poids = npoids
            qteColis = nqteColis
            bReturn = True
        Catch ex As Exception
            setError("CommandeClient.calcPoidsColis", ex.ToString)
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

    Public Function SetProperty(pKey As String, pValue As String) As Boolean
        Dim bReturn As Boolean
        Try
            Select Case pKey
                Case vncObjectProperties.CMD_ID
                    setid(CInt(pValue))
                    bReturn = True
                Case vncObjectProperties.CMD_CODE
                    code = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_dateCommande
                    dateCommande = ConvertJJMMAAAToDate(pValue)
                    bReturn = True
                Case vncObjectProperties.CMD_etat
                    setEtat(pValue)
                    bReturn = True
                Case vncObjectProperties.CMD_TotalHT
                    totalHT = CDec(pValue)
                    bReturn = True
                Case vncObjectProperties.CMD_TotalTTC
                    totalTTC = CDec(pValue)
                    bReturn = True

                Case vncObjectProperties.CMD_TransporteurCODE
                    Dim oParam As Transporteur
                    oParam = Transporteur.colTransporteur.Item(pValue)
                    If Not oParam Is Nothing Then
                        Me.oTransporteur = oParam
                        bReturn = True
                    Else
                        Trace.WriteLine("Transporteur code = " + pValue + " Non trouvé")
                        bReturn = False
                    End If
                Case vncObjectProperties.CMD_dateValidation
                    dateValidation = ConvertJJMMAAAToDate(pValue)
                    bReturn = True
                Case vncObjectProperties.CMD_dateEnlevement
                    dateEnlevement = ConvertJJMMAAAToDate(pValue)
                    bReturn = True
                Case vncObjectProperties.CMD_dateLivraison
                    dateLivraison = ConvertJJMMAAAToDate(pValue)
                    bReturn = True
                Case vncObjectProperties.CMD_RefLivraison
                    refLivraison = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TiersCODE
                    Dim oCol As Collection
                    Dim oTiers As Tiers
                    oTiers = Nothing
                    oCol = Client.getListe(pValue)
                    If Not oCol Is Nothing Then
                        If (oCol.Count = 1) Then
                            oTiers = oCol.Item(1)
                        End If
                    End If

                    If Not oTiers Is Nothing Then
                        oTiers.load()
                        Me.oTiers = oTiers
                        bReturn = True
                    Else
                        Trace.WriteLine("Tiers code = " + pValue + " Non trouvé")
                        bReturn = False
                    End If
                Case vncObjectProperties.CMD_TIERS_NOM
                    Me.caracteristiqueTiers.nom = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_RS
                    Me.caracteristiqueTiers.rs = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIV1
                    Me.caracteristiqueTiers.AdresseLivraison.rue1 = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIV2
                    Me.caracteristiqueTiers.AdresseLivraison.rue2 = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIVCP
                    Me.caracteristiqueTiers.AdresseLivraison.cp = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIVVILLE
                    Me.caracteristiqueTiers.AdresseLivraison.ville = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIVTEL
                    Me.caracteristiqueTiers.AdresseLivraison.tel = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIVFAX
                    Me.caracteristiqueTiers.AdresseLivraison.fax = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIVPORT
                    Me.caracteristiqueTiers.AdresseLivraison.port = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADLIVEMAIL
                    Me.caracteristiqueTiers.AdresseLivraison.Email = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACT1
                    Me.caracteristiqueTiers.AdresseFacturation.rue1 = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACT2
                    Me.caracteristiqueTiers.AdresseFacturation.rue2 = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACTCP
                    Me.caracteristiqueTiers.AdresseFacturation.cp = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACTVILLE
                    Me.caracteristiqueTiers.AdresseFacturation.ville = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACTTEL
                    Me.caracteristiqueTiers.AdresseFacturation.tel = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACTFAX
                    Me.caracteristiqueTiers.AdresseFacturation.fax = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACTPORT
                    Me.caracteristiqueTiers.AdresseFacturation.port = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_TIERS_ADFACTEMAIL
                    Me.caracteristiqueTiers.AdresseFacturation.Email = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_qteColis
                    Me.qteColis = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_qtePalettesNonPreparees
                    Me.qtePalettesNonPreparees = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_qtePalettesPreparees
                    Me.qtePalettesPreparees = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_poids
                    Me.poids = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_puPalettesNonPreparees
                    Me.puPalettesNonPreparees = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_puPalettesPreparees
                    Me.puPalettesPreparees = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_MontantTransport
                    Me.montantTransport = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_LettreVoiture
                    Me.lettreVoiture = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_CoutTransport
                    Me.coutTransport = CType(pValue, Decimal)
                    bReturn = True
                Case vncObjectProperties.CMD_RefFactTrp
                    Me.refFactTRP = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_ComentCom
                    Me.CommentaireCommandeText = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_ComentLiv
                    Me.CommentaireLivraisonText = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_ComentFact
                    Me.CommentaireFacturationText = pValue
                    bReturn = True
                Case vncObjectProperties.CMD_ComentLibre
                    Me.CommentaireLibreText = pValue
                    bReturn = True
                Case (vncObjectProperties.LGCMD_num)
                    'Recherche de la Ligne avec son Numéro
                    Dim oLg As LgCommande
                    Dim nValue As Integer
                    nValue = CType(pValue, Integer)
                    If Not bcolLignesLoaded Then
                        LoadcolLGCMD()
                    End If
                    m_oLgCourante = Nothing
                    For Each oLg In colLignes
                        If oLg.num = nValue Then
                            m_oLgCourante = oLg
                        End If
                    Next
                    If m_oLgCourante Is Nothing Then
                        m_oLgCourante = AjouteLigne() ' AjouteLigne d'une lign vide
                    Else
                        bReturn = True
                    End If

                Case vncObjectProperties.LGCMD_PRD_CODE
                    Dim oProduit As Produit
                    Dim oCol As Collection
                    oProduit = Nothing
                    oCol = Produit.getListe(vncEnums.vncTypeProduit.vncTous, pValue)
                    If oCol Is Nothing Then
                        Trace.WriteLine("produit code = " + pValue + " introuvable")
                        bReturn = False
                    Else
                        If oCol.Count <> 1 Then
                            Trace.WriteLine("produit code = " + pValue + " introuvable")
                            bReturn = False
                        Else
                            oProduit = CType(oCol(1), Produit)
                            If Not m_oLgCourante Is Nothing Then
                                m_oLgCourante.oProduit = oProduit
                                bReturn = True
                            Else
                                Trace.WriteLine("Pas de ligne courante")
                                bReturn = False
                            End If
                        End If
                    End If

                Case vncObjectProperties.LGCMD_qteCom
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.qteCommande = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_qteLiv
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.qteLiv = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_qteFact
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.qteFact = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_prixU
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.prixU = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_prixHT
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.prixHT = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_prixTTC
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.prixTTC = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_bGratuit
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.bGratuit = CType(pValue, Boolean)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_poids
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.poids = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If

                Case vncObjectProperties.LGCMD_qteColis
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.qteColis = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_TxComm
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.TxComm = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If
                Case vncObjectProperties.LGCMD_MtComm
                    If Not m_oLgCourante Is Nothing Then
                        m_oLgCourante.MtComm = CType(pValue, Decimal)
                        bReturn = True
                    Else
                        Trace.WriteLine("Pas de ligne courante")
                        bReturn = False
                    End If

            End Select
        Catch ex As InvalidCastException
            Trace.Write("Erreur en Conversion de " + pKey + "," + pValue)
            bReturn = False
        Catch ex As Exception
            Trace.Write(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SetProperty
    Public Function getProperty(pKey As String) As String
        Dim strReturn As String
        strReturn = ""
        Try
            Select Case pKey
                Case vncObjectProperties.CMD_ID
                    strReturn = Format(Me.id, "g")
                Case vncObjectProperties.CMD_CODE
                    strReturn = code
                Case vncObjectProperties.CMD_dateCommande
                    strReturn = Format(Me.dateCommande, "ddMMyyyy")
                Case vncObjectProperties.CMD_etat
                    strReturn = Me.etat.codeEtat
                Case vncObjectProperties.CMD_TotalHT
                    strReturn = Format(totalHT, "F")
                Case vncObjectProperties.CMD_TotalTTC
                    strReturn = Format(totalTTC, "F")

                Case vncObjectProperties.CMD_TransporteurCODE
                    strReturn = oTransporteur.code
                Case vncObjectProperties.CMD_dateValidation
                    strReturn = Format(dateValidation, "ddMMyyyy")
                Case vncObjectProperties.CMD_dateEnlevement
                    strReturn = Format(dateEnlevement, "ddMMyyyy")
                Case vncObjectProperties.CMD_dateLivraison
                    strReturn = Format(dateLivraison, "ddMMyyyy")
                Case vncObjectProperties.CMD_RefLivraison
                    strReturn = refLivraison
                Case vncObjectProperties.CMD_TiersCODE
                    strReturn = oTiers.code
                Case vncObjectProperties.CMD_TIERS_NOM
                    strReturn = Me.caracteristiqueTiers.nom
                Case vncObjectProperties.CMD_TIERS_RS
                    strReturn = Me.caracteristiqueTiers.rs
                Case vncObjectProperties.CMD_TIERS_ADLIV1
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.rue1
                Case vncObjectProperties.CMD_TIERS_ADLIV2
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.rue2
                Case vncObjectProperties.CMD_TIERS_ADLIVCP
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.cp

                Case vncObjectProperties.CMD_TIERS_ADLIVVILLE
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.ville

                Case vncObjectProperties.CMD_TIERS_ADLIVTEL
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.tel

                Case vncObjectProperties.CMD_TIERS_ADLIVFAX
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.fax

                Case vncObjectProperties.CMD_TIERS_ADLIVPORT
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.port

                Case vncObjectProperties.CMD_TIERS_ADLIVEMAIL
                    strReturn = Me.caracteristiqueTiers.AdresseLivraison.Email

                Case vncObjectProperties.CMD_TIERS_ADFACT1
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.rue1

                Case vncObjectProperties.CMD_TIERS_ADFACT2
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.rue2

                Case vncObjectProperties.CMD_TIERS_ADFACTCP
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.cp

                Case vncObjectProperties.CMD_TIERS_ADFACTVILLE
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.ville

                Case vncObjectProperties.CMD_TIERS_ADFACTTEL
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.tel

                Case vncObjectProperties.CMD_TIERS_ADFACTFAX
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.fax

                Case vncObjectProperties.CMD_TIERS_ADFACTPORT
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.port

                Case vncObjectProperties.CMD_TIERS_ADFACTEMAIL
                    strReturn = Me.caracteristiqueTiers.AdresseFacturation.Email

                Case vncObjectProperties.CMD_qteColis
                    strReturn = Format(Me.qteColis, "F")

                Case vncObjectProperties.CMD_qtePalettesNonPreparees
                    strReturn = Format(Me.qtePalettesNonPreparees, "F")

                Case vncObjectProperties.CMD_qtePalettesPreparees
                    strReturn = Format(Me.qtePalettesPreparees, "F")

                Case vncObjectProperties.CMD_poids
                    strReturn = Format(Me.poids, "F")

                Case vncObjectProperties.CMD_puPalettesNonPreparees
                    strReturn = Format(Me.puPalettesNonPreparees, "F")

                Case vncObjectProperties.CMD_puPalettesPreparees
                    strReturn = Format(Me.puPalettesPreparees, "F")

                Case vncObjectProperties.CMD_MontantTransport
                    strReturn = Format(Me.montantTransport, "F")

                Case vncObjectProperties.CMD_LettreVoiture
                    strReturn = Me.lettreVoiture

                Case vncObjectProperties.CMD_CoutTransport
                    strReturn = Format(Me.coutTransport, "F")

                Case vncObjectProperties.CMD_RefFactTrp
                    strReturn = Me.refFactTRP

                Case vncObjectProperties.CMD_ComentCom
                    strReturn = Me.CommentaireCommandeText
                Case vncObjectProperties.CMD_ComentLiv
                    strReturn = Me.CommentaireLivraisonText
                Case vncObjectProperties.CMD_ComentFact
                    strReturn = Me.CommentaireFacturationText
                Case vncObjectProperties.CMD_ComentLibre
                    strReturn = Me.CommentaireLibreText

                Case (vncObjectProperties.LGCMD_num)
                    strReturn = Format(m_oLgCourante.num, "g")
                Case vncObjectProperties.LGCMD_PRD_CODE
                    strReturn = m_oLgCourante.oProduit.code
                Case vncObjectProperties.LGCMD_qteCom
                    strReturn = Format(m_oLgCourante.qteCommande, "g")
                Case vncObjectProperties.LGCMD_qteLiv
                    strReturn = Format(m_oLgCourante.qteLiv, "g")
                Case vncObjectProperties.LGCMD_prixU
                    strReturn = Format(m_oLgCourante.prixU, "F")
                Case vncObjectProperties.LGCMD_prixHT
                    strReturn = Format(m_oLgCourante.prixHT, "F")
                Case vncObjectProperties.LGCMD_prixTTC
                    strReturn = Format(m_oLgCourante.prixTTC, "F")
                Case vncObjectProperties.LGCMD_bGratuit
                    strReturn = Format(m_oLgCourante.bGratuit, "TRUE/FALSE")
                Case vncObjectProperties.LGCMD_poids
                    strReturn = Format(m_oLgCourante.poids, "F")
                Case vncObjectProperties.LGCMD_qteColis
                    strReturn = Format(m_oLgCourante.qteColis, "F")
                Case vncObjectProperties.LGCMD_TxComm
                    strReturn = Format(m_oLgCourante.TxComm, "F")
                Case vncObjectProperties.LGCMD_MtComm
                    strReturn = Format(m_oLgCourante.MtComm, "F")

            End Select
        Catch ex As InvalidCastException
            Trace.Write("Erreur en Conversion de " + pKey)
            strReturn = ""
        Catch ex As Exception
            Trace.Write(ex.Message)
            strReturn = ""
        End Try
        Return strReturn
    End Function 'getProperty
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

    Public Shared Function importCommandeClient(pstrFileName As String) As CommandeClient
        Dim oReturn As Commande
        Dim oCmd As CommandeClient
        Dim oCLT As Client
        Dim oMyReader As FileIO.TextFieldParser
        Dim currentRow As String()
        Dim nField As Integer
        Dim oParam As ExportParam
        Dim oColExportParam As Collection
        Dim oLgCmd As LgCommande
        Try
            If Not My.Computer.FileSystem.FileExists(pstrFileName) Then
                Throw New Exception("Fichier d'entrée inexistant")
            End If
            'Lecture des paramètres d'importation
            oColExportParam = ExportParam.GetListe("CMDCLT")
            If oColExportParam Is Nothing Then
                Throw New Exception("No Param Found")
            End If
            If oColExportParam.Count = 0 Then
                Throw New Exception("No Param Found")
            End If
            'Création d'une commande vide
            oCLT = New Client()
            oCmd = New CommandeClient(oCLT)
            'Parcours du fichier texte (il concerne toute la commande)
            oMyReader = My.Computer.FileSystem.OpenTextFieldParser(pstrFileName, ";")
            nField = 0
            While Not oMyReader.EndOfData
                oLgCmd = oCmd.AjouteLigne()
                oCmd.setLgCourante(oLgCmd)
                currentRow = oMyReader.ReadFields()
                If (oColExportParam.Count < currentRow.Length) Then
                    Throw New Exception("Le nombre de champs dans le fichier texte dépasse le nombre de paramètre")
                End If
                For nField = 0 To currentRow.Length - 1
                    oParam = oColExportParam.Item(nField + 1)
                    If currentRow(nField).Length > 0 Then
                        oCmd.SetProperty(oParam.Valeur, currentRow(nField))
                    End If
                Next
                oCmd.m_oLgCourante.calcPoidsColis()
                oCmd.m_oLgCourante.CalculCommission()
                oCmd.m_oLgCourante.calculPrixTotal()
            End While

            oCmd.calculPrixTotal()
            oCmd.CalcMontantTransport()
            oMyReader.Close()
            oReturn = oCmd
        Catch ex As Exception
            Trace.WriteLine(ex.Message)
            oReturn = Nothing
        End Try
        Return oReturn
    End Function
End Class ' Commande
