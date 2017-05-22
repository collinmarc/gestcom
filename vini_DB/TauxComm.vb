'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : TauxComm 
' Description : Taux de commssion
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
'===================================================================================================================================
Imports System.IO
Public Class TauxComm
    Inherits Persist

    Private m_idFRN As Integer                 'ID Facture
    Private m_TypeClient As String               'Type de Facture
    Private m_Taux As Decimal                'Montant du reglement

#Region "Accesseurs"
    Public Sub New(ByVal pFRN As Fournisseur)
        Debug.Assert(pFRN.id > 0, "ID Fournisseur non renseigné")
        m_idFRN = pFRN.id
        m_TypeClient = Param.typeclientdefaut.code
        m_Taux = 0.0
    End Sub 'New
    Public Sub New(ByVal pFRNid As Integer)
        Debug.Assert(pFRNid > 0, "ID Fournisseur non renseigné")
        m_idFRN = pFRNid
        m_TypeClient = Param.typeclientdefaut.code
        m_Taux = 0.0
    End Sub 'New
    Public Sub New(ByVal pFRN As Fournisseur, ByVal pTypeClt As String, ByVal pTaux As Decimal)
        Debug.Assert(pFRN.id > 0, "ID Fournisseur non renseigné")
        m_idFRN = pFRN.id
        m_TypeClient = pTypeClt
        m_Taux = pTaux
    End Sub 'New
    Public Sub New(ByVal pFRNid As Integer, ByVal pTypeClt As String, ByVal pTaux As Decimal)
        Debug.Assert(pFRNid > 0, "ID Fournisseur non renseigné")
        m_idFRN = pFRNid
        m_TypeClient = pTypeClt
        m_Taux = pTaux
    End Sub 'New

    Public Sub New(ByVal pRow As System.Data.DataRow)
        Debug.Assert(pRow IsNot Nothing, "pRow non renseigné")
        getData(pRow)

    End Sub 'New
    Friend Function getData(ByVal pRow As System.Data.DataRow) As Boolean

        Dim oRow As dsVinicom.FRN_COMMRow
        Dim bReturn As Boolean

        Try
            oRow = CType(pRow, dsVinicom.FRN_COMMRow)
            setid(oRow.FRNC_ID)
            m_idFRN = oRow.FRNC_FRN_ID
            m_TypeClient = oRow.FRNC_TYPE_CLIENT.Trim()
            If Not oRow.IsFRNC_TX_COMMNull() Then
                m_Taux = oRow.FRNC_TX_COMM
            End If
            resetBooleans()
            bReturn = True

        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function



    Public Property IdFRN() As Integer
        Get
            Return m_idFRN
        End Get
        Set(ByVal Value As Integer)
            If Not m_idFRN.Equals(Value) Then
                m_idFRN = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property TypeClt() As String
        Get
            Return m_TypeClient
        End Get
        Set(ByVal Value As String)
            If Not m_TypeClient.Equals(Value) Then
                m_TypeClient = Value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property TauxComm() As Decimal
        Get
            Return m_Taux
        End Get
        Set(ByVal Value As Decimal)
            If Value <> m_Taux Then
                m_Taux = Value
                RaiseUpdated()
            End If
        End Set
    End Property

#End Region
#Region "Interface Persist"
    '=======================================================================
    'Fonction : DBLoad()
    'Description : Chargement de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If

        Debug.Assert(id <> 0, "idCommande <> 0")
        bReturn = loadTauxComm()
        Return bReturn
    End Function 'DBLoad
    Public Overrides Function save() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        bReturn = MyBase.Save()
        shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : delete()
    'Description : Suppression de l'objet dans la base de l'objet
    'Détails    :  
    'Retour : Vrai si l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Debug.Assert(id <> 0, "idCommande <> 0")
        Dim bReturn As Boolean
        bReturn = deleteTauxComm()
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
        Debug.Assert(m_idFRN <> 0, "L'Id Fournisseur doit être Renseigné")
        Debug.Assert(Not String.IsNullOrEmpty(m_TypeClient), "Le type de client doit être Renseigné")

        Dim bReturn As Boolean
        bReturn = insertTauxComm()
        Return bReturn
    End Function 'insert
    '=======================================================================
    'Fonction : Update()
    'Description : Mise à jour de l'objet
    'Détails    :  
    'Retour : Vrai di l'opération s'est bien déroulée
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Debug.Assert(m_idFRN <> 0, "L'Id Fournisseur doit être Renseigné")
        Debug.Assert(String.IsNullOrEmpty(m_TypeClient), "Le type de client doit être Renseigné")
        Debug.Assert(id <> 0, "id <> 0")
        Dim bReturn As Boolean
        bReturn = updateTauxComm()
        Return bReturn

    End Function 'Update

    ''' <summary>
    ''' Rend la Liste des Taux de Commission associé à une fournisseur
    ''' </summary>
    ''' <param name="pIdFRN">id Fournisseur</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getListe(ByVal pIdFRN As Integer) As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = getListeTauxComm(pIdFRN)
        shared_disconnect()
        Return colReturn
    End Function
    ''' <summary>
    ''' Rend la Liste des Taux de Commission associé à un fournisseur et un type de Client
    ''' </summary>
    ''' <param name="pIdFRN">id Fournisseur</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getListe(ByVal pIdFRN As Integer, ByVal pTypeclient As String) As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = getListeTauxComm(pIdFRN, pTypeclient)
        Debug.Assert(colReturn.Count <= 1, "Plus d'un Element retourné dans Tauxcomm.getListe")

        shared_disconnect()
        Return colReturn
    End Function
#End Region

    '=======================================================================
    'Fonction : shortResume()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "TauxComm(" & m_id & "," & m_idFRN & "," & m_TypeClient & "," & m_Taux.ToString & ")"
        End Get
    End Property

    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "TauxComm(" & m_id & "," & m_idFRN & "," & m_TypeClient & "," & m_Taux.ToString & ")"
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim obj1 As TauxComm
        Try

            bReturn = obj.GetType.Name.Equals(Me.GetType().Name)
            obj1 = obj
            bReturn = MyBase.Equals(obj)

            bReturn = bReturn And (m_idFRN.Equals(obj1.IdFRN))
            bReturn = bReturn And (m_TypeClient.Equals(obj1.TypeClt))
            bReturn = bReturn And (m_Taux.Equals(obj1.TauxComm))

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn
    End Function 'Equals

End Class
