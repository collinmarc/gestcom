'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Param
' Description : Paramètre
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
'===================================================================================================================================Public MustInherit Class Persist
Public Class Param
    Inherits Persist
    Private m_type As String
    Private m_code As String
    Private m_valeur As String
    Private m_valeur2 As String
    Private m_defaut As Boolean
    Private Shared m_colConstantes As New Collection
    Private Shared m_colModeReglement As New Collection
    Private Shared m_colTypeClient As New Collection
    Private Shared m_colCouleur As New Collection
    Private Shared m_colRegion As New Collection
    Private Shared m_colConditionnement As New Collection
    Private Shared m_colTVA As New Collection
    Private Shared m_colTypeCommande As New Collection
    Private Shared m_colTransport As New Collection
    Private Shared m_bcolparamsLoad As Boolean = False
    Private Shared m_CouleurDefaut As New Param
    Private Shared m_ModeReglementDefaut As New Param
    Private Shared m_TypeClientDefaut As New Param
    Private Shared m_RegionDefaut As New Param
    Private Shared m_ConditionnementDefaut As New Param
    Private Shared m_TVADefaut As New Param

    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : LoadcolParams 
    'Description : Chargement des paramètres
    'Retour :  
    '=======================================================================
    Public Shared Sub LoadcolParams()
        Dim oParam As Param
        Persist.shared_connect()
        m_colModeReglement = Persist.getListeParam(PAR_MODE_RGLMT)
        For Each oParam In m_colModeReglement
            If oParam.m_defaut Then
                m_ModeReglementDefaut = oParam
                Exit For
            End If
        Next
        m_colTypeClient = Persist.getListeParam(PAR_TYPE_CLIENT)
        For Each oParam In m_colTypeClient
            If oParam.m_defaut Then
                m_TypeClientDefaut = oParam
                Exit For
            End If
        Next
        m_colCouleur = Persist.getListeParam(PAR_COULEUR)
        For Each oParam In m_colCouleur
            If oParam.m_defaut Then
                m_CouleurDefaut = oParam
                Exit For
            End If
        Next
        m_colRegion = Persist.getListeParam(PAR_REGION)
        For Each oParam In m_colRegion
            If oParam.m_defaut Then
                m_RegionDefaut = oParam
                Exit For
            End If
        Next
        m_colConditionnement = Persist.getListeParam(PAR_CONDITIONNEMENT)
        For Each oParam In m_colConditionnement
            If oParam.m_defaut Then
                m_ConditionnementDefaut = oParam
                Exit For
            End If
        Next
        m_colTVA = Persist.getListeParam(PAR_TVA)
        For Each oParam In m_colTVA
            If oParam.m_defaut Then
                m_TVADefaut = oParam
                Exit For
            End If
        Next

        m_colConstantes = Persist.getListeConstantes()

        m_bcolparamsLoad = True
        Persist.shared_disconnect()
    End Sub 'LoadColParams
    Public Shared ReadOnly Property colModeReglement() As Collection
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_colModeReglement
        End Get
    End Property
    Public Shared ReadOnly Property colTypeClient() As Collection
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_colTypeClient
        End Get
    End Property
    Public Shared ReadOnly Property colCouleur() As Collection
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_colCouleur
        End Get
    End Property
    Public Shared ReadOnly Property colRegion() As Collection
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_colRegion
        End Get
    End Property
    Public Shared ReadOnly Property colConditionnement() As Collection
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_colConditionnement
        End Get
    End Property
    Public Shared ReadOnly Property colTVA() As Collection
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_colTVA
        End Get
    End Property
    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : couleurDefaut()
    'Description : Rend le paramétre par defaut
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared ReadOnly Property couleurdefaut() As Param
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_CouleurDefaut
        End Get
    End Property
    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : ModeReglementdefaut()
    'Description : Rend le paramétre par defaut
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared ReadOnly Property ModeReglementdefaut() As Param
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_ModeReglementDefaut
        End Get
    End Property
    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : typeclientdefaut()
    'Description : Rend le paramétre par defaut
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared ReadOnly Property typeclientdefaut() As Param
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_TypeClientDefaut
        End Get
    End Property

    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : regiondefaut()
    'Description : Rend le paramétre par defaut
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared ReadOnly Property regiondefaut() As Param
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_RegionDefaut
        End Get
    End Property
    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : conditionnementdefaut()
    'Description : Rend le paramétre par defaut
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared ReadOnly Property conditionnementdefaut() As Param
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_ConditionnementDefaut
        End Get
    End Property

    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : TVAdefaut()
    'Description : Rend le paramétre par defaut
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared ReadOnly Property TVAdefaut() As Param
        Get
            If Not m_bcolparamsLoad Then
                LoadcolParams()
            End If
            Return m_TVADefaut
        End Get
    End Property
    '=======================================================================
    '               MEMBRE DE CLASSE
    'Fonction : getConstante(strNom)
    'Description : Rend la constante indiquée ou ""
    'Détails    :  
    'Retour : un Paramétre
    '=======================================================================
    Public Shared Function getConstante(ByVal strNom As String) As String
        If Not m_bcolparamsLoad Then
            LoadcolParams()
        End If
        Try
            Return m_colConstantes(strNom)
        Catch ex As Exception
            Return ""
        End Try
    End Function


    Public Property type() As String

        Get
            Return m_type
        End Get
        Set(ByVal Value As String)
            If Not m_type.Equals(Value) Or Value Is Nothing Then
                m_type = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property code() As String
        Get
            Return m_code
        End Get
        Set(ByVal Value As String)
            If Not m_code.Equals(Value) Or Value Is Nothing Then
                m_code = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property valeur() As String
        Get
            Return m_valeur
        End Get
        Set(ByVal Value As String)
            If Not m_valeur.Equals(Value) Or Value Is Nothing Then
                m_valeur = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property valeur2() As String
        Get
            Return m_valeur2
        End Get
        Set(ByVal Value As String)
            If Not m_valeur2.Equals(Value) Or Value Is Nothing Then
                m_valeur2 = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property defaut() As Boolean
        Get
            Return m_defaut
        End Get
        Set(ByVal Value As Boolean)
            If Not m_defaut.Equals(Value) Then
                m_defaut = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Overrides Function toString() As String
        Return "PAR =(" & m_type & "," & m_code & "," & m_valeur & "," & m_defaut & ")"
    End Function
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objP As Param
        Dim bReturn As Boolean
        Try
            bReturn = True
            objP = CType(obj, Param)

            bReturn = bReturn And (id.Equals(objP.id))
            bReturn = bReturn And (code.Equals(objP.code))
            bReturn = bReturn And (type.Equals(objP.type))
            bReturn = bReturn And (valeur.Equals(objP.valeur))
            bReturn = bReturn And (defaut.Equals(objP.defaut))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub New()
        m_type = String.Empty
        m_code = String.Empty
        m_valeur = String.Empty
        m_valeur2 = String.Empty
        m_defaut = False
        m_id = 0
    End Sub

    Public Sub New(ByVal pType As String)
        m_type = pType
        m_code = String.Empty
        m_valeur = String.Empty
        m_valeur2 = String.Empty
        m_defaut = False
        m_id = 0
    End Sub
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_code
        End Get
    End Property
    Public Shared ReadOnly Property LGNUM_GAZOLE() As String
        Get
            Return CST_LGFACTTRP_NUM_GO
        End Get
    End Property

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        bReturn = loadParam()
        Return bReturn

    End Function


    '=======================================================================
    'Fonction : delete
    'Description : suppression de l'objet en base
    'Détails    : Appelle deletePRD (de Persist) 
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = deleteParam()
        shared_disconnect()
        Return bReturn
    End Function 'delete
    '=======================================================================
    'Fonction : checkFordelete
    'description : Controle si l'élément est supprimable
    '                table commandesClients
    '=======================================================================
    Public Overrides Function checkForDelete() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = False
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete
    '=======================================================================
    'Fonction : Insert
    'Description : Insert de l'objet en base
    'Détails    : Appelle InsertPRD (de Persist) 
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = insertParam()
        shared_disconnect()
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : update
    'Description : Update de l'objet en base
    'Détails    : Appelle UpdatePRD (de Persist) 
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = updateParam()
        shared_disconnect()
        Return bReturn
    End Function 'update

    Friend Overrides Function fromRS(ByVal poRS As OleDb.OleDbDataReader) As Boolean
        Dim bReturn As Boolean
        Try

            MyBase.fromRS(poRS)

            Try
                Me.setid(GetString(poRS, "PAR_ID"))
            Catch ex As InvalidCastException
                Me.setid(0)
            End Try
            Try
                Me.code = GetString(poRS, "PAR_CODE")
            Catch ex As InvalidCastException
                Me.code = ""
            End Try
            Try
                Me.type = GetString(poRS, "PAR_TYPE")
            Catch ex As InvalidCastException
                Me.type = ""
            End Try
            Try
                Me.valeur = GetString(poRS, "PAR_VALUE")
            Catch ex As InvalidCastException
                Me.valeur = ""
            End Try
            Try
                Me.valeur2 = GetString(poRS, "PAR_VALUE2")
            Catch ex As InvalidCastException
                Me.valeur2 = ""
            End Try
            Try
                Me.defaut = GetString(poRS, "PAR_DEFAUT")
            Catch ex As InvalidCastException
                Me.defaut = False
            End Try

            bReturn = True
        Catch ex As Exception
            setError(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
End Class
