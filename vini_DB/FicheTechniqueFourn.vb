Imports System.Drawing
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 06/12/11
'===================================================================================================================================
' Classe : FicheTechniqueFourn
' Description : Fichier technique Fournisseur
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
Public Class FicheTechniqueFourn
    Inherits Persist
    Private m_TxtPresDomaine As String 'Présentation du domaine
    Private m_TxtPresFournisseur As String 'Présentation du fournisseur
    Private m_Specialite As String 'Specialité (séparée par des |)
    Private m_txtTerroir As String 'Descrption du terroir
    Private m_IMGDomaine As vini_Image
    Private m_IMGFournisseur As vini_Image


    Public Sub New(ByVal pfrnId As Long)
        setid(pfrnId)
        txtPresDomaine = ""
        specialite = ""
        txtPresFournisseur = ""
        txtTerroir = ""
        m_IMGDomaine = New vini_Image(pfrnId, "FTFRN_DOM", 1)
        m_IMGFournisseur = New vini_Image(pfrnId, "FTFRN_FRN", 1)
    End Sub

    Public Property txtPresDomaine As String
        Get
            Return m_TxtPresDomaine
        End Get
        Set(ByVal Value As String)
            m_TxtPresDomaine = Value
            RaiseUpdated()
        End Set

    End Property
    Public Property txtPresFournisseur As String
        Get
            Return m_TxtPresFournisseur
        End Get
        Set(ByVal Value As String)
            m_TxtPresFournisseur = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property specialite As String
        Get
            Return m_Specialite
        End Get
        Set(ByVal Value As String)
            m_Specialite = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property txtTerroir As String
        Get
            Return m_txtTerroir
        End Get
        Set(ByVal Value As String)
            m_txtTerroir = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property IMGDomaine As vini_Image
        Get
            Return m_IMGDomaine
        End Get
        Set(ByVal Value As vini_Image)
            m_IMGDomaine = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property IMGFournisseur As vini_Image
        Get
            Return m_IMGFournisseur
        End Get
        Set(ByVal Value As vini_Image)
            m_IMGFournisseur = Value
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
        Dim strReturn As String
        strReturn = "FICHETECHNIQUEFOUN =( frnID" & m_id & ")\n "
        strReturn = strReturn & "txtPresDomaine =<" & txtPresDomaine & ">\n "
        strReturn = strReturn & "txtPresFournisseur=<" & txtPresFournisseur & ">\n "
        strReturn = strReturn & "specialite=<" & specialite & ">\n "
        strReturn = strReturn & "txtTerroir=<" & txtTerroir & ">\n "
        Return strReturn
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim objFT As FicheTechniqueFourn
        Dim bReturn As Boolean
        Try
            bReturn = True
            objFT = CType(obj, FicheTechniqueFourn)

            bReturn = MyBase.Equals(obj)
            bReturn = bReturn And (objFT.id.Equals(Me.id))
            bReturn = bReturn And (objFT.txtPresDomaine.Equals(Me.txtPresDomaine))
            bReturn = bReturn And (objFT.txtPresFournisseur.Equals(Me.txtPresFournisseur))
            bReturn = bReturn And (objFT.specialite.Equals(Me.specialite))
            bReturn = bReturn And (objFT.txtTerroir.Equals(Me.txtTerroir))

            Return bReturn
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function


    '=======================================================================
    'Fonction : shortResume()
    'Description : Rend un resumé de l'objet sous forme de chaine
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides ReadOnly Property shortResume() As String
        Get
            Return "FTFRN"
        End Get
    End Property

    Public Overrides Function checkForDelete() As Boolean
        Return True
    End Function

    Protected Overrides Function DBLoad(Optional pid As Integer = 0) As Boolean
        Dim bReturn As Boolean
        If pid <> 0 Then
            m_id = pid
        End If
        bReturn = loadFicheTechniqueFourn()
        If bReturn Then
            bReturn = IMGDomaine.LoadFromObject(m_id, "FTFRN_DOM", 1)
        End If
        If bReturn Then
            bReturn = IMGFournisseur.LoadFromObject(m_id, "FTFRN_FRN", 1)
        End If
        Return bReturn
    End Function

    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        bReturn = deleteFicheTechniqueFourn()
        Return bReturn
    End Function

    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        bReturn = insertFicheTechniqueFourn()
        If (bReturn) Then
            bReturn = saveImages()
        End If
        Return bReturn
    End Function

    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        bReturn = updateFicheTechniqueFourn()
        If (bReturn) Then
            bReturn = saveImages()
        End If
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : SaveImages
    'Description : Sauvegardes des images d'une fiche technique founisseur
    'Retour : Rend Vrai si l'INSERT s'est correctement effectué
    '=======================================================================
    Private Function saveImages() As Boolean
        Dim bReturn As Boolean

        Try
            bReturn = False
            IMGDomaine.IdObject = id
            bReturn = IMGDomaine.Save()
            If bReturn Then
                IMGFournisseur.IdObject = id
                bReturn = IMGFournisseur.Save()
            End If

        Catch ex As Exception
            setError("FicheTechniqueFournisseur.SaveImages", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SaveImages

    Friend Function getData(ByVal pRow As System.Data.DataRow) As Boolean

        Dim oRow As dsVinicom.FICHETECHNIQUE_FOURNISSEURRow
        Dim bReturn As Boolean

        Try
            oRow = CType(pRow, dsVinicom.FICHETECHNIQUE_FOURNISSEURRow)
            setid(oRow.FTFRN_ID)
            If Not oRow.IsFTFRN_TXTDOMAINENull Then
                txtPresDomaine = oRow.FTFRN_TXTDOMAINE
            End If
            If Not oRow.IsFTFRN_TXTFOURNISSEURNull() Then
                txtPresFournisseur = oRow.FTFRN_TXTFOURNISSEUR
            End If

            If Not oRow.IsFTFRN_TXTTERROIRNull() Then
                txtTerroir = oRow.FTFRN_TXTTERROIR
            End If
            If Not oRow.IsFTFRN_SPECIALITENull() Then
                specialite = oRow.FTFRN_SPECIALITE
            End If
            bReturn = True

        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function
    Public Function loadIMGDomaine(FilePath As String) As Boolean

        Dim bReturn As Boolean
        Try
            bReturn = False
            If System.IO.File.Exists(FilePath) Then
                m_IMGDomaine.LoadFromFile(FilePath)
                RaiseUpdated()
                bReturn = True
            End If
        Catch ex As Exception
            setError("FicheTechniqueFourn.loadIMGDomaine", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
    Public Function loadIMGFournisseur(FilePath As String) As Boolean

        Dim bReturn As Boolean
        Try
            bReturn = False
            If System.IO.File.Exists(FilePath) Then
                m_IMGFournisseur.LoadFromFile(FilePath)
                RaiseUpdated()
                bReturn = True
            End If
        Catch ex As Exception
            setError("FicheTechniqueFourn.loadIMGFournisseur", ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function
End Class
