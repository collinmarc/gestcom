'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : lgFactTRP
' Description : Ligne de Facture de transport
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
'
'===================================================================================================================================
Public Class LgFactTRP
    Inherits Persist

    Private m_num As Integer                'Numéro d'ordre de ligne
    Private m_idFactTRP As Long
    Private m_idCmdCLT As Long
    Private m_Montant_HT As Decimal
    Private m_nomTransporteur As String
    Private m_dateLivraison As Date
    Private m_referenceLivraison As String
    Private m_dateCommande As Date
    Private m_refCommande As String
    Private m_qteColis As Decimal
    Private m_qtePalettesPreparees As Decimal
    Private m_qtePalettesNonPreparees As Decimal
    Private m_poids As Decimal
    Private m_puPalettesPreparees As Decimal
    Private m_puPalettesNonPreparees As Decimal



#Region "Accesseurs"
    Public Property num() As Integer
        Get
            Return m_num
        End Get
        Set(ByVal Value As Integer)
            If m_num <> Value Then
                m_num = Value
                bUpdated = True
            End If
        End Set
    End Property
    Public Sub New()
        m_typedonnee = vncEnums.vncTypeDonnee.LGFACTTRP
        m_idFactTRP = 0
        m_idCmdCLT = 0
        m_Montant_HT = 0
        m_nomTransporteur = ""
        m_dateLivraison = DATE_DEFAUT
        m_referenceLivraison = ""
        m_dateCommande = DATE_DEFAUT
        m_refCommande = ""
        m_qteColis = 0
        m_qtePalettesPreparees = 0
        m_qtePalettesNonPreparees = 0
        m_poids = 0
        m_puPalettesPreparees = 0
        m_puPalettesNonPreparees = 0
    End Sub
    Public Property idFactTRP() As Long
        Get
            Return m_idFactTRP
        End Get
        Set(ByVal Value As Long)
            If Value <> m_idFactTRP Then
                m_idFactTRP = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property idCmdCLT() As Long
        Get
            Return m_idCmdCLT
        End Get
        Set(ByVal Value As Long)
            If m_idCmdCLT <> Value Then
                m_idCmdCLT = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property prixHT() As Decimal
        Get
            Return m_Montant_HT
        End Get
        Set(ByVal Value As Decimal)
            If m_Montant_HT <> Value Then
                m_Montant_HT = Value
                RaiseUpdated()
            End If
        End Set
    End Property


    Public Property nomTransporteur() As String
        Get
            Return m_nomTransporteur
        End Get
        Set(ByVal Value As String)
            If (Not Value.Equals(m_nomTransporteur)) Then
                m_nomTransporteur = Value
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
    Public Property referenceLivraison() As String
        Get
            Return m_referenceLivraison
        End Get
        Set(ByVal Value As String)
            If (Not Value.Equals(m_referenceLivraison)) Then
                m_referenceLivraison = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property dateCommande() As Date
        Get
            Return m_dateCommande
        End Get
        Set(ByVal Value As Date)
            If m_dateCommande <> Value Then
                m_dateCommande = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property refCommande() As String
        Get
            Return m_refCommande
        End Get
        Set(ByVal Value As String)
            If (Not Value.Equals(m_refCommande)) Then
                m_refCommande = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qteColis() As Decimal
        Get
            Return m_qteColis
        End Get
        Set(ByVal Value As Decimal)
            If m_qteColis <> Value Then
                m_qteColis = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qtePalettesPreparees() As Decimal
        Get
            Return m_qtePalettesPreparees
        End Get
        Set(ByVal Value As Decimal)
            If m_qtePalettesPreparees <> Value Then
                m_qtePalettesPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property qtePalettesNonPreparees() As Decimal
        Get
            Return m_qtePalettesNonPreparees
        End Get
        Set(ByVal Value As Decimal)
            If m_qtePalettesNonPreparees <> Value Then
                m_qtePalettesNonPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property puPalettesPreparees() As Decimal
        Get
            Return m_puPalettesPreparees
        End Get
        Set(ByVal Value As Decimal)
            If m_puPalettesPreparees <> Value Then
                m_puPalettesPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property puPalettesNonPreparees() As Decimal
        Get
            Return m_puPalettesNonPreparees
        End Get
        Set(ByVal Value As Decimal)
            If m_puPalettesNonPreparees <> Value Then
                m_puPalettesNonPreparees = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Property poids() As Decimal
        Get
            Return m_poids
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_poids Then
                m_poids = Value
                RaiseUpdated()
            End If
        End Set
    End Property


#End Region
#Region "Méthodes Persist"

    '=======================================================================
    'Fonction : DBLoad()
    'Description : Chargement de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Return bReturn
    End Function 'DBLoad

    '=======================================================================
    'Fonction : delete()
    'Description : Suppression de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai si l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Return bReturn

    End Function ' delete
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

    '=======================================================================
    'Fonction : insert()
    'Description : insertion de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean

        bReturn = False
        Return bReturn

    End Function 'insert

    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Return bReturn

    End Function 'Update
#End Region

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "LGFACTTRP =(" & m_id & "," & m_idFactTRP & "," & m_idCmdCLT & "," & m_Montant_HT & "," & m_nomTransporteur & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objLgFactTRP As LgFactTRP
        Try
            objLgFactTRP = CType(obj, LgFactTRP)
            bReturn = True
            bReturn = bReturn And (m_num.Equals(objLgFactTRP.num))
            bReturn = bReturn And (m_idCmdCLT = objLgFactTRP.idCmdCLT)
            bReturn = bReturn And (m_idFactTRP = objLgFactTRP.idFactTRP)
            bReturn = bReturn And (m_Montant_HT = objLgFactTRP.prixHT)
            bReturn = bReturn And (m_nomTransporteur.Equals(objLgFactTRP.nomTransporteur))
            bReturn = bReturn And (m_referenceLivraison.Equals(objLgFactTRP.referenceLivraison))
            bReturn = bReturn And (m_dateCommande.Equals(objLgFactTRP.dateCommande))
            bReturn = bReturn And (m_refCommande.Equals(objLgFactTRP.refCommande))
            bReturn = bReturn And (m_qteColis = objLgFactTRP.qteColis)
            bReturn = bReturn And (m_qtePalettesPreparees = objLgFactTRP.qtePalettesPreparees)
            bReturn = bReturn And (m_qtePalettesNonPreparees = objLgFactTRP.qtePalettesNonPreparees)
            bReturn = bReturn And (m_poids = objLgFactTRP.poids)
            bReturn = bReturn And (m_puPalettesPreparees = objLgFactTRP.puPalettesPreparees)
            bReturn = bReturn And (m_puPalettesNonPreparees = objLgFactTRP.puPalettesNonPreparees)
        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'Equals

    '=======================================================================
    'Fonction : calculPrixTotal()
    'Description : Calcul du prix total d'une lligne de transport
    'Détails    :  Met a jour l'attribut PrixHT
    'Retour : TRUE/FALSE
    '=======================================================================
    Public Function calculPrixTotal() As Boolean
        Dim bReturn As Boolean
        Try
            prixHT = (qtePalettesPreparees * puPalettesPreparees) + (qtePalettesNonPreparees * puPalettesNonPreparees)
            bReturn = True
        Catch ex As Exception
            setError("LgFactTRP.calculPrixTotal", ex.ToString)
            bReturn = False
        End Try
    End Function

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "LGTRP"
        End Get
    End Property
    '===================================
    ' Function : dupliqueinfosCommande
    ' Description : initialise une ligne de transport avec une Commande Client
    '===================================
    Public Function dupliqueinfosCommande(ByVal objCmd As CommandeClient) As Boolean
        Debug.Assert(objCmd.id <> 0, "Id de la commande <> 0")
        Debug.Assert(objCmd.etat.codeEtat <> vncEnums.vncEtatCommande.vncLivree Or _
            objCmd.etat.codeEtat <> vncEnums.vncEtatCommande.vncTransmise Or _
            objCmd.etat.codeEtat <> vncEnums.vncEtatCommande.vncEclatee Or _
            objCmd.etat.codeEtat <> vncEnums.vncEtatCommande.vncRapprochee _
            , "Etat de la commande incorect" & objCmd.etat.libelle)

        Dim bReturn As Boolean = False
        Try
            bReturn = True
            Debug.Assert(Not objCmd Is Nothing)
            idCmdCLT = objCmd.id
            prixHT = Math.Round(objCmd.montantTransport, 2)
            nomTransporteur = objCmd.oTransporteur.nom
            dateLivraison = objCmd.dateLivraison
            referenceLivraison = objCmd.refLivraison
            dateCommande = objCmd.dateCommande
            refCommande = objCmd.code
            qteColis = objCmd.qteColis
            qtePalettesPreparees = objCmd.qtePalettesPreparees
            qtePalettesNonPreparees = objCmd.qtePalettesNonPreparees
            puPalettesPreparees = objCmd.puPalettesPreparees
            puPalettesNonPreparees = objCmd.puPalettesNonPreparees
            poids = objCmd.poids
            calculPrixTotal()


        Catch ex As Exception
            bReturn = False
            setError("LgFactTRP.duppliqueInfosCommande", ex.ToString)
        End Try

        Debug.Assert(bReturn, getErreur())
        Return bReturn
    End Function 'DuppliqueInfosCommande
End Class
