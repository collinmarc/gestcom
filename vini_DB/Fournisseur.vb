Option Explicit On 
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 23/06/04
'===================================================================================================================================
' Classe : Fournisseur
' Description : Contient les informations du fournisseur
'===================================================================================================================================
'Membres de Classes
'==================
'Public
'   getListe    : Rend la Liste des Fournisseurs
'Protected
'Private
'Membres d'instances
'==================
'Public
'   idRegion : Identifiant de la région Vinicole
'   libRegion : (ReadOnly) Libellé de la région Vinicole
'   load(id) : Chargement de l'objet en base
'               Rend True si le chargement est OK
'   insert(id) : insertio de l'objet en base
'               Rend True si l'insertion  est OK
'   update(id) : MAJ de l'objet en base
'               Rend True si la MAJ est OK
'   delete(id) : Suppression de l'objet en base
'               Rend True si le DELETE est OK
'Protected
'Private
'===================================================================================================================================
'Historique :
'============
'
'===================================================================================================================================
Public Class Fournisseur
    Inherits Tiers
    Private m_idRegion As Integer
    Private m_libRegion As String
    Private m_bExportInternet As Integer
    '=======================================================================
    '                           METHODE DE CLASSE                          |  
    'Fonction : getListe 
    'Description : Liste des Fournisseurs
    'Retour : Rend une collection de fournisseurs
    '=======================================================================
    Public Shared Function getListe(Optional ByVal strCode As String = "", Optional ByVal strNom As String = "", Optional ByVal strRs As String = "") As Collection
        Dim colReturn As Collection
        shared_connect()
        colReturn = ListeFRN(strCode, strNom, strRS)
        shared_disconnect()
        Return colReturn
    End Function
    'Identifiant de la région
    Public Property idRegion() As Integer
        Get
            Return m_idRegion
        End Get
        Set(ByVal Value As Integer)
            RaiseUpdated()
            m_idRegion = Value
        End Set
    End Property
    'Libellé de la région vinicole
    Public Property libregion() As String
        Get
            Return m_libRegion
        End Get
        Set(ByVal Value As String)
            m_libRegion = Value
        End Set

    End Property
    'Export des bons  à facture sur internet
    Public Property bExportInternet() As Integer
        Get
            Debug.Assert(Not m_bResume, "Objet de type Résumé")
            Return m_bExportInternet
        End Get
        Set(ByVal Value As Integer)
            RaiseUpdated()
            m_bExportInternet = Value
        End Set

    End Property
    Public Shared Function createandload(ByVal pid As Integer) As Fournisseur
        '=======================================================================
        ' Contructeur pour chargement
        '=======================================================================
        Dim objFourn As Fournisseur
        Dim bReturn As Boolean
        objFourn = New Fournisseur
        Try
            If pid <> 0 Then
                bReturn = objFourn.load(pid)
                If Not bReturn Then
                    setError("Fournisseur.createAndLoad", getErreur())
                End If

            End If
        Catch ex As Exception
            setError("Fournisseur.createAndLoad", ex.ToString)
        End Try
        Debug.Assert(objFourn.id = pid, "Fournisseur " & pid & " non chargée")
        Return objFourn
    End Function 'Createanload
    ''' <summary>
    ''' Constructeur pour Chargement par la clé
    ''' </summary>
    ''' <param name="pCode"> Code Fournisseur</param>
    ''' <returns>Objet Fournisseur ou null</returns>
    ''' <remarks></remarks>
    Public Shared Function createandload(ByVal pCode As String) As Fournisseur
        Dim objFourn As Fournisseur
        Dim bReturn As Boolean
        Dim nId As Integer
        objFourn = New Fournisseur
        Try
            If Not String.IsNullOrEmpty(pCode) Then
                nId = Fournisseur.getFRNIDByKey(pCode)
                If nId <> -1 Then
                    bReturn = objFourn.load(nId)
                    If Not bReturn Then
                        setError("Fournisseur.createAndLoad", getErreur())
                        objFourn = Nothing
                    End If
                Else
                    setError("Fournisseur.createAndLoad", "No ID for " & pCode)
                    objFourn = Nothing
                End If
            End If
        Catch ex As Exception
            setError("Fournisseur.createAndLoad", ex.ToString)
            objFourn = Nothing
        End Try
        Return objFourn
    End Function 'Createanload

    '=======================================================================
    'Fonction : Load
    'Description : Chargement de l'objet en base
    'Détails    : Appelle LoadFRN (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Protected Overrides Function DBLoad(Optional ByVal pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        shared_connect()
        If pid <> 0 Then
            m_id = pid
        End If
        bReturn = loadFRN()
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            m_bNew = False
            setUpdatedFalse()
            m_bDeleted = False
        End If
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : LoadLight
    'Description : Chargement de l'objet en base
    'Détails    : Appelle LoadFRN (de Persist) 
    'Retour : Rend Vrai si le chargement s'est correctement effectué
    '=======================================================================
    Friend Overrides Function LoadLight() As Boolean
        Debug.Assert(m_id <> 0, "L'id doit être initialisé")

        Dim bReturn As Boolean
        shared_connect()
        bReturn = loadFRNLight()
        shared_disconnect()
        'Mise à jour des indicateurs
        If bReturn Then
            m_bNew = False
            setUpdatedFalse()
            m_bDeleted = False
        End If
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : delete
    'Description : suppression de l'objet en base
    'Détails    : Appelle deleteFRN (de Persist) 
    'Retour : Rend Vrai si le DELETE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        shared_connect()
        TauxComm.deleteTauxComms(m_id)
        bReturn = deleteFRN()
        shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : checkFordelete
    'description : Controle si l'élément est supprimable
    '                table commandesClients
    '=======================================================================
    Public Overrides Function checkForDelete() As Boolean
        Dim bReturn As Boolean
        Try
            '    shared_connect()
            bReturn = True
            If existeProduitFournisseur() Then
                bReturn = False
            End If
            If bReturn And existeBonApproFournisseur() Then
                bReturn = False
            End If
            '   shared_disconnect()
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function 'checkForDelete
    '=======================================================================
    'Fonction : Insert
    'Description : Insert de l'objet en base
    'Détails    : Appelle InsertFRN (de Persist) 
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        '    shared_connect()
        bReturn = InsertFRN()
        '   shared_disconnect()
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : update
    'Description : Update de l'objet en base
    'Détails    : Appelle UpdateFRN (de Persist) 
    'Retour : Rend Vrai si l'UPDATE s'est correctement effectué
    '=======================================================================
    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        '  shared_connect()
        bReturn = updateFRN()
        ' shared_disconnect()
        Return bReturn
    End Function

    '=======================================================================
    'Fonction : Equals
    'Description : Comparaison d'objet
    'Détails    : Compare chaque membre de l'objet 
    'Retour : Rend Vrai si l'objet passé en paramétre est équivalent
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objFrn As Fournisseur
        Dim bReturn As Boolean
        Try
            bReturn = True
            objFrn = CType(obj, Fournisseur)

            bReturn = MyBase.Equals(obj)
            bReturn = bReturn And (objFrn.idRegion.Equals(Me.idRegion))
            bReturn = bReturn And (objFrn.bExportInternet.Equals(Me.bExportInternet))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    '=======================================================================
    'Fonction : new
    'Description : Création de l'objet
    'Détails    : 
    'Retour : 
    '=======================================================================

    Public Sub New(ByVal pcode As String, ByVal pnom As String)
        MyBase.New(pcode, pnom)
        m_typedonnee = vncEnums.vncTypeDonnee.FOURNISSEUR
        m_idRegion = Param.regiondefaut.id
        m_libRegion = Param.regiondefaut.valeur
        m_bExportInternet = 1
    End Sub
    Public Sub New()
        MyBase.New("", "")
        m_typedonnee = vncEnums.vncTypeDonnee.FOURNISSEUR
        m_idRegion = Param.regiondefaut.id
        m_libRegion = Param.regiondefaut.valeur
        m_bExportInternet = 1
    End Sub

    '=======================================================================
    'Fonction : toString
    'Description : Rend L'objet sous forme de chaine
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Return "FRN : [" & MyBase.toString() & "]" & idRegion & "," & libregion
    End Function

    Public Sub createTxCommStandard(pTypeClient As String)
        Try

            Dim oTx As New TauxComm(Me.id, pTypeClient, Param.getConstante("CST_TX_COMMISSION"))
            oTx.save()
        Catch ex As Exception

        End Try
    End Sub

End Class
