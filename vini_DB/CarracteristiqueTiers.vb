'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Tiers
' Description : Objet CaracteristiqueTiers : Contient les adresses et les commentaires d'un tiers
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'Protected
'Private
'Membres d'instances
'==================
'Public
'   Code            : Code Tiers
'   nom             : Nom
'   rs              : Raison Sociale
'   motcle          : Mot cle de reherche
'   rib1            : Code Banque
'   rib2            : Code Guichet
'   rib3            : Numéro de compte
'   rib4            : Clé RIB
'   banque          : Banque de Reglement
'   siret           : Numéro de SIRET
'   tvaintracom     : Numéro de TVA intraCommunautaire
'   idModeRglmt     : Id du mode de reglement
'   libModeRglmt    : Libellé du mode de reglement
'   colAdresse      : Collection d'adresses
'   colComment      : Collection de Commentaires
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================Public MustInherit Class Persist
Public MustInherit Class CaracteristiqueTiers
    Inherits Persist
    Private m_code As String            'code du tiers
    Private m_nom As String             'Nom du tiers
    Private m_rs As String              ' Raison Sociale
    Private m_rib1 As String            'Code Banque
    Private m_rib2 As String            'Code Guichet
    Private m_rib3 As String            'Numéro de compte
    Private m_rib4 As String            ' Clé RIB
    Private m_banque As String          'Banque de Reglement
    Private m_siret As String           'Numéro de SIRET
    Private m_tvaintracom As String     ' Numéro de TVA intraCommunautaire
    Private m_idModeRglmt As Integer    'Id du mode de reglement
    Private m_libModeRglmt As String    'Libellé du mode de reglement
    Private WithEvents m_colAdresse As ColEvent ' Collection d'adresses
    Private WithEvents m_colComment As ColEvent ' Collection de Commentaires
    Private m_bAdressesIdentiques As Boolean 'Adresse de Facturation et de livraison identiques
    Private m_Codecompta As String
    Private m_idModeRglmt1 As Integer    'Id du mode de reglement Facture
    Private m_idModeRglmt2 As Integer    'Id du mode de reglement Facture 2
    Private m_idModeRglmt3 As Integer    'Id du mode de reglement Facture 3


    Friend MustOverride Function loadLight() As Boolean
    ' Public MustOverride Function delete() As Boolean
    ' Public MustOverride Function insert() As Boolean
    ' Public MustOverride Function Load(ByVal pid As Integer) As Boolean
    ' Public MustOverride Function update() As Boolean


    Public Property code() As String
        Get
            Return m_code
        End Get
        Set(ByVal Value As String)
            If Value <> m_code Then
                RaiseUpdated()
                m_code = Value
            End If
        End Set
    End Property
    Public Property nom() As String
        Get
            Return m_nom
        End Get
        Set(ByVal Value As String)
            If Value <> m_nom Then
                RaiseUpdated()
                m_nom = Value
            End If
        End Set
    End Property
    Public Property rs() As String
        Get
            Return m_rs
        End Get
        Set(ByVal Value As String)
            If Value <> m_rs Then
                RaiseUpdated()
                m_rs = Value
            End If
        End Set
    End Property
    Public Property rib1() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_rib1
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_rib1 <> Value Then
                    RaiseUpdated()
                    m_rib1 = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property
    Public Property rib2() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_rib2
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_rib2 <> Value Then
                    RaiseUpdated()
                    m_rib2 = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property
    Public Property rib3() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_rib3
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_rib3 <> Value Then
                    RaiseUpdated()
                    m_rib3 = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property
    Public Property rib4() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_rib4
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_rib4 <> Value Then
                    RaiseUpdated()
                    m_rib4 = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property
    Public Property banque() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_banque
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_banque <> Value Then
                    RaiseUpdated()
                    m_banque = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property

    Public Property siret() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_siret
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_siret <> Value Then
                    RaiseUpdated()
                    m_siret = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property

    Public Property numtvaintracom() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_tvaintracom
        End Get
        Set(ByVal Value As String)
            If Not bResume Then
                If m_tvaintracom <> Value Then
                    RaiseUpdated()
                    m_tvaintracom = Value
                End If
            Else
                Throw New System.Exception("Propriété non modifiable dans un objet Résumé")
            End If
        End Set
    End Property

    Public Property bAdressesIdentiques() As Boolean
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_bAdressesIdentiques
        End Get
        Set(ByVal Value As Boolean)
            If m_bAdressesIdentiques <> Value Then
                m_bAdressesIdentiques = Value
                RaiseUpdated()
            End If
        End Set
    End Property

    Public Sub New(ByVal pCode As String, ByVal pnom As String)

        Dim objAdrLiv As Adresse
        Dim objAdrFact As Adresse
        Dim objComment As Commentaire

        m_code = pCode
        m_nom = pnom
        m_rs = ""
        m_rib1 = ""
        m_rib2 = ""
        m_rib3 = ""
        m_rib4 = ""
        m_banque = ""
        m_siret = ""
        m_tvaintracom = ""
        m_idModeRglmt = Param.ModeReglementdefaut.id
        m_libModeRglmt = Param.ModeReglementdefaut.valeur
        m_idModeRglmt1 = Param.ModeReglementdefaut.id
        m_idModeRglmt2 = Param.ModeReglementdefaut.id
        m_idModeRglmt3 = Param.ModeReglementdefaut.id
        m_libModeRglmt = Param.ModeReglementdefaut.valeur
        'Création de la collection des adresses
        m_colAdresse = New ColEvent
        objAdrLiv = New Adresse
        objAdrLiv.Code = ADRLIV
        m_colAdresse.Add(objAdrLiv, objAdrLiv.Code)
        objAdrFact = New Adresse
        objAdrFact.Code = ADRFACT
        m_colAdresse.Add(objAdrFact, objAdrFact.Code)
        'Création de la collection des commentaires
        m_colComment = New ColEvent
        objComment = New Commentaire
        objComment.code = COMCMD
        m_colComment.Add(objComment, objComment.code)
        objComment = New Commentaire
        objComment.code = COMLIV
        m_colComment.Add(objComment, objComment.code)
        objComment = New Commentaire
        objComment.code = COMFACT
        m_colComment.Add(objComment, objComment.code)
        objComment = New Commentaire
        objComment.code = COMLIBRE
        m_colComment.Add(objComment, objComment.code)
        m_bAdressesIdentiques = False
        m_Codecompta = String.Empty
    End Sub
    'Commentaires
    Public ReadOnly Property colCommentaires() As ColEvent
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colComment
        End Get
    End Property
    'Commentaire de commande
    Public ReadOnly Property CommCommande() As Commentaire
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colComment(COMCMD)
        End Get
    End Property
    'Commentaire de commande
    Public ReadOnly Property CommLivraison() As Commentaire
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colComment(COMLIV)
        End Get
    End Property
    'Commentaire de commande
    Public ReadOnly Property CommFacturation() As Commentaire
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colComment(COMFACT)
        End Get
    End Property
    'Commentaire libre
    Public ReadOnly Property CommLibre() As Commentaire
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colComment(COMLIBRE)
        End Get
    End Property

    'Adresses
    Public ReadOnly Property colAdresses() As ColEvent
        Get
            Return m_colAdresse
        End Get
    End Property

    Public ReadOnly Property AdresseLivraison() As Adresse
        Get
            Return m_colAdresse(ADRLIV)
        End Get
    End Property

    Public ReadOnly Property AdresseLivraisonNom() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.nom
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonRue1() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.rue1
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonRue2() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.rue2
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonCP() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.cp
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonTel() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.tel
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonVille() As String
        Get
            Dim oAdresse As Adresse
            Try
                oAdresse = m_colAdresse(ADRLIV)
                Return oAdresse.ville
            Catch ex As Exception
                Return String.Empty
            End Try
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonFax() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.fax
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonPort() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.port
        End Get
    End Property
    Public ReadOnly Property AdresseLivraisonEmail() As String
        Get
            Dim oAdresse As Adresse
            oAdresse = m_colAdresse(ADRLIV)
            Return oAdresse.Email
        End Get
    End Property
    Public ReadOnly Property AdresseFacturation() As Adresse
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_colAdresse(ADRFACT)
        End Get
    End Property

    'Mode de reglement
    Public Property idModeReglement() As Integer
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_idModeRglmt
        End Get
        Set(ByVal Value As Integer)
            RaiseUpdated()
            m_idModeRglmt = Value
        End Set
    End Property

    Public Property libModeReglement() As String
        Get
            Debug.Assert(Not m_bResume, "Objet de type resumé")
            Return m_libModeRglmt
        End Get
        Set(ByVal Value As String)
            m_libModeRglmt = Value
        End Set
    End Property
    Public Property CodeCompta() As String
        Get
            Return m_Codecompta
        End Get
        Set(ByVal Value As String)
            m_Codecompta = Value
        End Set
    End Property
    'Mode de reglement de facture 1
    Public Property idModeReglement1() As Integer
        Get
            Return m_idModeRglmt1
        End Get
        Set(ByVal Value As Integer)
            If Not m_idModeRglmt1.Equals(Value) Then
                m_idModeRglmt1 = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    'Mode de reglement de facture 2
    Public Property idModeReglement2() As Integer
        Get
            Return m_idModeRglmt2
        End Get
        Set(ByVal Value As Integer)
            If Not m_idModeRglmt2.Equals(Value) Then
                m_idModeRglmt2 = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    'Mode de reglement de facture 3
    Public Property idModeReglement3() As Integer
        Get
            Return m_idModeRglmt3
        End Get
        Set(ByVal Value As Integer)
            If Not m_idModeRglmt3.Equals(m_idModeRglmt) Then
                m_idModeRglmt3 = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Overrides Function toString() As String
        Dim sREturn As String
        sREturn = "Tiers : (" & MyBase.toString() & m_code & "," & m_nom & "," & m_rs & "," & m_rib1 & "," & m_rib2 & "," & m_rib3 & "," & m_rib4 & "," & m_banque & "," & m_siret & "," & m_tvaintracom & "," & m_idModeRglmt & "," & m_libModeRglmt & "," & bAdressesIdentiques
        sREturn = sREturn & m_colAdresse.toString()
        sREturn = sREturn & m_colComment.toString()
        sREturn = sREturn & ")"
        Return sREturn
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objTiers As Tiers
        Dim objAdresse As Adresse
        Dim objComment As Commentaire
        Dim bReturn As Boolean
        Try
            objTiers = CType(obj, Tiers)
            bReturn = MyBase.Equals(obj)
            bReturn = bReturn And (m_nom.Equals(objTiers.nom))
            bReturn = bReturn And (m_code.Equals(objTiers.code))
            bReturn = bReturn And (m_rs.Equals(objTiers.rs))
            bReturn = bReturn And (m_rib1.Equals(objTiers.rib1))
            bReturn = bReturn And (m_rib2.Equals(objTiers.rib2))
            bReturn = bReturn And (m_rib3.Equals(objTiers.rib3))
            bReturn = bReturn And (m_rib4.Equals(objTiers.rib4))
            bReturn = bReturn And (m_banque.Equals(objTiers.banque))
            bReturn = bReturn And (m_siret.Equals(objTiers.siret))
            bReturn = bReturn And (m_tvaintracom.Equals(objTiers.numtvaintracom))
            bReturn = bReturn And (m_idModeRglmt.Equals(objTiers.idModeReglement))
            bReturn = bReturn And (m_bAdressesIdentiques.Equals(objTiers.bAdressesIdentiques))
            For Each objAdresse In m_colAdresse
                '                objAdresse = DE.Value
                Try
                    bReturn = bReturn And (objAdresse.Equals(objTiers.colAdresses(objAdresse.Code)))
                Catch ex As Exception
                    bReturn = False
                    Exit For
                End Try

            Next
            For Each objComment In m_colComment
                Try
                    bReturn = bReturn And (objComment.Equals(objTiers.colCommentaires(objComment.code)))
                Catch ex As Exception
                    bReturn = False
                    Exit For
                End Try

            Next
            'bReturn = bReturn And (m_colAdresse.Equals(objTiers.colAdresses))
            'bReturn = bReturn And (m_colComment.Equals(objTiers.colCommentaires))
            bReturn = bReturn And (m_Codecompta.Equals(objTiers.CodeCompta))


            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Sub m_colAdresse_Updated() Handles m_colAdresse.Updated
        raiseupdated()
    End Sub

    Private Sub m_colComment_Updated() Handles m_colComment.Updated
        raiseUpdated()
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_code & " | " & m_rs & " | " & m_nom & " | " & AdresseLivraison.ville
        End Get
    End Property


End Class
