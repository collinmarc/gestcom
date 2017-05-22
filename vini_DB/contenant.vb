'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : conditionnement
' Description : Classe de description des conditionnement
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
'  17/10/05 : Ajout du poids
'
'===================================================================================================================================Public MustInherit Class Persist
Public Class contenant
    Inherits Persist


    Private m_code As String
    Private m_libelle As String
    Private m_cent As Double
    Private m_bout As Double
    Private m_poids As Decimal
    Private m_defaut As Boolean
    Private Shared m_colContenant As New Collection
    Private Shared m_bcolLoad As Boolean = False
    Private Shared m_contenantDefaut As New contenant

    '=======================================================================
    '                   MEMBRE DE CLASSE
    'Fonction : colcolntenant
    'Description : Liste des contenant de l'application
    'Détails    : Appelle LoadcolParam 
    '=======================================================================
    Public Shared ReadOnly Property colContenant() As Collection
        Get
            If Not m_bcolLoad Then
                LoadcolContenants()
                m_bcolLoad = True
            End If
            Return m_colContenant
        End Get
    End Property
    '=======================================================================
    '                   METHODE DE CLASSE
    'Fonction : LoadcolParam
    'Description : Chargement de la liste des contenants et intialise le contenant par defaut
    'Détails    : Appelle getListContenant (de Persist) 
    'Retour : Rien
    '=======================================================================
    Public Shared Sub LoadcolContenants()
        Dim oCont As contenant
        Dim bConnectionOpen As Boolean
        bConnectionOpen = Persist.shared_isConnected()
        If Not bConnectionOpen Then
            Persist.shared_connect()

        End If
        m_colContenant = Persist.getListeContenants()
        For Each oCont In m_colContenant
            If oCont.defaut Then
                m_contenantDefaut = oCont
                Exit For
            End If
        Next
        If Not bConnectionOpen Then
            Persist.shared_disconnect()
        End If

    End Sub
    '=======================================================================
    '                   METHODE DE CLASSE
    'Fonction : ContenantDefaut
    'Description : Rend le contenant par defaut
    'Détails    : loadcolContenants
    'Retour : Rien
    '=======================================================================
    Public Shared ReadOnly Property contenantDefaut() As contenant
        Get
            If Not m_bcolLoad Then
                LoadcolContenants()
            End If
            Return m_contenantDefaut
        End Get
    End Property

    Public Property code() As String
        Get
            Return m_code
        End Get
        Set(ByVal Value As String)
            m_code = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property libelle() As String
        Get
            Return m_libelle
        End Get
        Set(ByVal Value As String)
            m_libelle = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property cent() As Double
        Get
            Return m_cent
        End Get
        Set(ByVal Value As Double)
            m_cent = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property bout() As Double
        Get
            Return m_bout
        End Get
        Set(ByVal Value As Double)
            m_bout = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property defaut() As Boolean
        Get
            Return m_defaut
        End Get
        Set(ByVal Value As Boolean)
            m_defaut = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property poids() As Decimal
        Get
            Return m_poids
        End Get
        Set(ByVal Value As Decimal)
            m_poids = Value
            RaiseUpdated()
        End Set
    End Property

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "CONT =(" & m_id & "," & m_code & "," & m_libelle & "," & m_cent & "," & m_bout & ")"
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objCont As contenant
        Dim bReturn As Boolean
        Try
            bReturn = True
            objCont = CType(obj, contenant)

            bReturn = bReturn And (id.Equals(objCont.id))
            bReturn = bReturn And (code.Equals(objCont.code))
            bReturn = bReturn And (libelle.Equals(objCont.libelle))
            bReturn = bReturn And (cent.Equals(objCont.cent))
            bReturn = bReturn And (bout.Equals(objCont.bout))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub New()
        m_code = ""
        m_libelle = ""
        m_defaut = False
        m_bout = 1
        m_cent = 1
    End Sub

    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return m_code
        End Get
    End Property

    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        Debug.Assert(id <> 0, "idCommande <> 0")
        bReturn = loadContenant()
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
        bReturn = deleteContenant()
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
        bReturn = insertContenant()
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
        bReturn = updateContenant()
        shared_disconnect()
        Return bReturn
    End Function 'update

    Friend Overrides Function fromRS(ByVal poRS As OleDb.OleDbDataReader) As Boolean
        Dim bReturn As Boolean
        Try

            MyBase.fromRS(poRS)

            Try
                setid(getInteger(poRS, "CONT_ID"))
            Catch ex As InvalidCastException
                Me.setid(0)
            End Try
            Try
                code = GetString(poRS, "CONT_CODE")
            Catch ex As InvalidCastException
                Me.code = ""
            End Try
            libelle = GetString(poRS, "CONT_LIBELLE")
            cent = GetString(poRS, "CONT_CENT")
            bout = GetString(poRS, "CONT_BOUT")
            poids = GetString(poRS, "CONT_POIDS")
            defaut = GetString(poRS, "CONT_DEFAUT")
            bReturn = True
        Catch ex As Exception
            setError(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function

End Class
