Imports System.Drawing
Imports System.IO
'===================================================================================================================================
'Projet : Vinicom
'Auteur : Marc Collin 
'Création : 06/12/11
'===================================================================================================================================
' Classe : vini_image
' Description : Une imahe du fichier technique
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
Public Class vini_Image
    Inherits Persist
    Private m_IdObject As Integer   'Identifiant de l'objet associé
    Private m_type As String   'Type de l'image
    Private m_Num As Integer   'Numéro de l'image
    Private m_Image As Byte() 'Image
    Private m_Desc As String 'Description

#Region "Constructeur"
    Public Sub New()
    End Sub
    Public Sub New(pIdObject As Integer, ptype As String, pNum As Integer)
        m_IdObject = pIdObject
        m_type = ptype
        m_Num = pNum
        m_Desc = ""
    End Sub
#End Region
#Region "Properties"
    Public Property IdObject As Integer
        Get
            Return m_IdObject
        End Get
        Set(ByVal Value As Integer)
            m_IdObject = Value
            RaiseUpdated()
        End Set

    End Property
    Public Property Num As Integer
        Get
            Return m_Num
        End Get
        Set(ByVal Value As Integer)
            m_Num = Value
            RaiseUpdated()
        End Set

    End Property
    Public Property Type As String
        Get
            Return m_type
        End Get
        Set(ByVal Value As String)
            m_type = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property Image As Byte()
        Get
            Return m_Image
        End Get
        Set(ByVal Value As Byte())
            m_Image = Value
            RaiseUpdated()
        End Set
    End Property
    Public Property Desc As String
        Get
            Return m_Desc
        End Get
        Set(ByVal Value As String)
            m_Desc = Value
            RaiseUpdated()
        End Set
    End Property
#End Region
    '=======================================================================
    'Fonction : ToString()
    'Description : Rend l'objet sous forme de String
    'Détails    :  
    'Retour : une Chaine
    '=======================================================================
    Public Overrides Function toString() As String
        Dim strReturn As String
        strReturn = "VINI_IMAGE =(ID=" & m_id & ",m_objectID =" & m_IdObject & ","
        strReturn = strReturn & "type=" & Type & ","
        strReturn = strReturn & "num=" & Num & ","
        strReturn = strReturn & "desc=" & Desc & ""
        Return strReturn
    End Function

    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal pobj As Object) As Boolean
        Dim obj As vini_image
        Dim bReturn As Boolean
        Try
            bReturn = True
            obj = CType(pobj, vini_image)

            bReturn = MyBase.Equals(obj)
            bReturn = bReturn And (obj.IdObject.Equals(Me.IdObject))
            bReturn = bReturn And (obj.Type.Equals(Me.Type))
            bReturn = bReturn And (obj.Num.Equals(Me.Num))
            If Me.Image Is Nothing Then
                bReturn = obj.Image Is Nothing
            Else
                bReturn = bReturn And (obj.Image.Equals(Me.Image))
            End If
            If Me.Desc Is Nothing Then
                bReturn = obj.Desc Is Nothing
            Else
                bReturn = bReturn And (obj.Desc.Equals(Me.Desc))
            End If

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
            Return "vini_image"
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
        bReturn = loadImage()
        Return bReturn
    End Function

    Public Function LoadFromObject(pObjectid As Integer, pType As String, pNum As Integer) As Boolean
        Dim bReturn As Boolean
        bReturn = loadImage(pObjectid, pType, pNum)
        Return bReturn
    End Function
    Friend Overrides Function delete() As Boolean
        Dim bReturn As Boolean
        bReturn = deleteImage()
        Return bReturn
    End Function

    Friend Overrides Function insert() As Boolean
        Dim bReturn As Boolean
        bReturn = insertImage()
        Return bReturn
    End Function

    Friend Overrides Function update() As Boolean
        Dim bReturn As Boolean
        bReturn = updateImage()
        Return bReturn
    End Function

    Friend Function getData(ByVal pRow As System.Data.DataRow) As Boolean

        Dim oRow As dsVinicom.IMAGESRow
        Dim bReturn As Boolean

        Try
            oRow = CType(pRow, dsVinicom.IMAGESRow)
            setid(oRow.IMG_ID)
            If Not oRow.IsIMG_OBJECT_IDNull Then
                IdObject = oRow.IMG_OBJECT_ID
            End If
            If Not oRow.IsIMG_TYPENull() Then
                Type = oRow.IMG_TYPE
            End If

            If Not oRow.IsIMG_NUMNull() Then
                Num = oRow.IMG_NUM
            End If
            If Not oRow.IsIMG_IMAGENull() Then
                Image = oRow.IMG_IMAGE
            End If
            If Not oRow.IsIMG_DESCNull() Then
                Desc = oRow.IMG_DESC
            End If
            bReturn = True

        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            bReturn = False
        End Try

        Return bReturn
    End Function ' GetData

    Public Function LoadFromFile(pstrFilePath As String) As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Try
            If File.Exists(pstrFilePath) Then
                Image = File.ReadAllBytes(pstrFilePath)
                bReturn = True
            End If
        Catch ex As Exception
            setError(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'LoadFromFile
    Public Function SaveToFile(pstrFilePath As String) As Boolean
        Dim bReturn As Boolean
        bReturn = False
        Try
            If Image.Length > 0 Then
                If File.Exists(pstrFilePath) Then
                    File.Delete(pstrFilePath)
                End If
                File.WriteAllBytes(pstrFilePath, Image)
                bReturn = True
            End If

        Catch ex As Exception
            setError(ex.Message)
            bReturn = False
        End Try
        Return bReturn
    End Function 'SaveToFile
End Class
